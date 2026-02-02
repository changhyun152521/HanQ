"""
학생 선택 모달(출제 대상 선택)

요구사항:
- 출제하기 클릭 시 학생 선택 모달 표시
- 학년/반 탭 제공 (반은 아직 로드 기능 미구현 → UI만 유지)
- 검색 기능
- 여러 학생 다중 선택
- 학년 전체 선택 가능(학년 그룹 체크박스로 처리)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Set

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
)

from core.models import Student
from database.repositories.class_repository import ClassRepository
from database.repositories.student_repository import StudentRepository


def _font(size_pt: int, *, bold: bool = False, extra_bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    if extra_bold:
        f.setWeight(QFont.ExtraBold)
    elif bold:
        f.setBold(True)
    else:
        try:
            f.setWeight(QFont.Medium)
        except Exception:
            pass
    return f


def _grade_rank(grade: str) -> int:
    g = (grade or "").strip()
    if not g:
        return 9_999
    base = 9_000
    if g.startswith("초"):
        base = 0
    elif g.startswith("중"):
        base = 100
    elif g.startswith("고"):
        base = 200
    n = ""
    for ch in g:
        if ch.isdigit():
            n += ch
    try:
        return base + int(n or 99)
    except Exception:
        return base + 99


@dataclass
class SelectedStudents:
    ids: List[str]


# 체크박스는 보통 행 왼쪽 고정 너비 안에 그려짐 (스타일 API 의존 없이 판별)
_CHECKBOX_ZONE_WIDTH = 28


def _is_checkbox_click(tree: QTreeWidget, item: QTreeWidgetItem, event_pos) -> bool:
    """클릭 위치가 해당 행의 체크박스 영역인지 여부. 체크박스가 있는 행만 True 가능."""
    if (item.flags() & Qt.ItemIsUserCheckable) == 0:
        return False
    idx = tree.indexFromItem(item)
    item_rect = tree.visualRect(idx)
    pos_in_vp = tree.viewport().mapFrom(tree, event_pos)
    if not item_rect.contains(pos_in_vp):
        return False
    # 행 안에서 왼쪽 체크박스 구역(픽셀) 안이면 체크박스 클릭으로 간주
    return (pos_in_vp.x() - item_rect.x()) < _CHECKBOX_ZONE_WIDTH


class _StudentTreeWidget(QTreeWidget):
    """학년/반 그룹: 체크박스 클릭 = 전체 선택/해제, 행(나머지) 클릭 = 펼치기/접기."""

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            item = self.itemAt(event.pos())
            if item is not None and item.childCount() > 0:
                if _is_checkbox_click(self, item, event.pos()):
                    # 체크박스 클릭 → 기본 동작(체크 토글 → _on_item_changed에서 전체 선택 처리)
                    super().mousePressEvent(event)
                    return
                # 행(체크박스 제외) 클릭 → 드롭다운만
                idx = self.indexFromItem(item)
                self.setExpanded(idx, not self.isExpanded(idx))
                event.accept()
                return
        super().mousePressEvent(event)


class StudentSelectDialog(QDialog):
    def __init__(self, *, db_connection, parent=None):
        super().__init__(parent)
        self.db_connection = db_connection
        self.repo: Optional[StudentRepository] = None
        self.class_repo: Optional[ClassRepository] = None
        try:
            if self.db_connection is not None and getattr(self.db_connection, "is_connected", None) and self.db_connection.is_connected():
                self.repo = StudentRepository(self.db_connection)
                self.class_repo = ClassRepository(self.db_connection)
        except Exception:
            self.repo = None
            self.class_repo = None

        self._mode: str = "grade"  # grade | class
        self._students: List[Student] = []
        self._classes: List = []  # List[SchoolClass]
        self._selected_ids: Set[str] = set()
        self._building = False

        self._tree: Optional[QTreeWidget] = None
        self._search: Optional[QLineEdit] = None
        self._lbl_count: Optional[QLabel] = None
        self._info: Optional[QLabel] = None

        self.setWindowTitle("학생 선택")
        self.setModal(True)
        self.setMinimumWidth(520)
        self.setMinimumHeight(560)

        self._build_ui()
        self._load_students()
        self._rebuild()

    def selected_student_ids(self) -> List[str]:
        return sorted(list(self._selected_ids))

    def _build_ui(self) -> None:
        self.setObjectName("StudentSelectDialog")
        self.setStyleSheet(
            """
            QDialog#StudentSelectDialog {
                background: #FFFFFF;
            }

            QFrame#SegWrap {
                background: #F1F5F9;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
            }
            QPushButton#SegBtn {
                border: none;
                border-radius: 12px;
                padding: 10px 0px;
                color: #334155;
                background: transparent;
                font-weight: 800;
            }
            QPushButton#SegBtn:checked {
                background: #2563EB;
                color: #FFFFFF;
                font-weight: 900;
            }

            QLineEdit#SearchBox {
                background-color: #FFFFFF;
                border: 1.5px solid #475569;
                border-radius: 12px;
                padding: 10px 12px;
                color: #0F172A;
                font-weight: 800;
            }
            QLineEdit#SearchBox:focus {
                border: 2px solid #2563EB;
            }
            QLineEdit#SearchBox::placeholder {
                color: #64748B;
            }

            QTreeWidget#StudentTree {
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                padding: 6px;
                background: #FFFFFF;
            }
            QTreeWidget#StudentTree::item {
                padding: 8px 6px;
                color: #0F172A;
            }
            QTreeWidget#StudentTree::item:hover {
                background: #F8FAFC;
                border-radius: 10px;
            }

            QPushButton#PrimaryBtn {
                background: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 900;
            }
            QPushButton#PrimaryBtn:hover {
                background: #1D4ED8;
            }
            QPushButton#GhostBtn {
                background: #F1F5F9;
                color: #1E293B;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 800;
            }
            QPushButton#GhostBtn:hover {
                background: #E2E8F0;
            }
            """
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(12)

        title = QLabel("출제할 학생을 선택하세요")
        title.setFont(_font(12, extra_bold=True))
        title.setStyleSheet("color:#0F172A;")
        root.addWidget(title)

        # 세그먼트(학년/반)
        seg = QFrame()
        seg.setObjectName("SegWrap")
        seg_lay = QHBoxLayout(seg)
        seg_lay.setContentsMargins(2, 2, 2, 2)
        seg_lay.setSpacing(2)

        self.btn_grade = QPushButton("학년")
        self.btn_grade.setObjectName("SegBtn")
        self.btn_grade.setCheckable(True)
        self.btn_grade.setChecked(True)
        self.btn_grade.setFocusPolicy(Qt.NoFocus)
        self.btn_grade.setCursor(Qt.PointingHandCursor)

        self.btn_class = QPushButton("반")
        self.btn_class.setObjectName("SegBtn")
        self.btn_class.setCheckable(True)
        self.btn_class.setChecked(False)
        self.btn_class.setFocusPolicy(Qt.NoFocus)
        self.btn_class.setCursor(Qt.PointingHandCursor)

        self.btn_grade.clicked.connect(lambda: self._set_mode("grade"))
        self.btn_class.clicked.connect(lambda: self._set_mode("class"))

        seg_lay.addWidget(self.btn_grade, 1)
        seg_lay.addWidget(self.btn_class, 1)
        root.addWidget(seg)

        # 검색
        self._search = QLineEdit()
        self._search.setObjectName("SearchBox")
        self._search.setPlaceholderText("학생 이름 검색")
        self._search.setClearButtonEnabled(True)
        self._search.setFont(_font(10, bold=True))
        self._search.textChanged.connect(self._rebuild)
        root.addWidget(self._search)

        # 안내
        self._info = QLabel("")
        self._info.setFont(_font(9, bold=True))
        self._info.setStyleSheet("color:#64748B;")
        self._info.setWordWrap(True)
        root.addWidget(self._info)

        # 트리 (학년/반 그룹 행 셀 전체 클릭 시에도 펼치기/접기)
        tree = _StudentTreeWidget()
        self._tree = tree
        tree.setObjectName("StudentTree")
        tree.setHeaderHidden(True)
        tree.setColumnCount(1)
        tree.itemChanged.connect(self._on_item_changed)
        root.addWidget(tree, 1)

        # 하단 바
        bottom = QHBoxLayout()
        bottom.setContentsMargins(0, 0, 0, 0)
        bottom.setSpacing(10)

        self._lbl_count = QLabel("선택 0명")
        self._lbl_count.setFont(_font(9, bold=True))
        self._lbl_count.setStyleSheet("color:#475569;")
        bottom.addWidget(self._lbl_count, alignment=Qt.AlignVCenter)

        bottom.addStretch(1)

        btn_cancel = QPushButton("취소")
        btn_cancel.setObjectName("GhostBtn")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setFocusPolicy(Qt.NoFocus)
        btn_cancel.clicked.connect(self.reject)
        bottom.addWidget(btn_cancel)

        btn_ok = QPushButton("출제하기")
        btn_ok.setObjectName("PrimaryBtn")
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setFocusPolicy(Qt.NoFocus)
        btn_ok.clicked.connect(self._on_ok)
        bottom.addWidget(btn_ok)

        root.addLayout(bottom)

    def _set_mode(self, mode: str) -> None:
        m = "class" if (mode or "").strip() == "class" else "grade"
        self._mode = m
        self.btn_grade.setChecked(m == "grade")
        self.btn_class.setChecked(m == "class")
        if m == "class":
            self._load_classes()
        self._rebuild()

    def _load_students(self) -> None:
        self._students = []
        if self.repo is None:
            return
        try:
            self._students = self.repo.list_all()
        except Exception:
            self._students = []

    def _load_classes(self) -> None:
        self._classes = []
        if self.class_repo is None:
            return
        try:
            self._classes = self.class_repo.list_all()
        except Exception:
            self._classes = []

    def _filtered_students(self) -> List[Student]:
        q = ""
        if self._search is not None:
            q = (self._search.text() or "").strip().lower()
        if not q:
            return list(self._students)
        out: List[Student] = []
        for s in self._students:
            if q in ((s.name or "").lower()):
                out.append(s)
        return out

    def _rebuild(self) -> None:
        tree = self._tree
        if tree is None:
            return

        self._building = True
        try:
            tree.blockSignals(True)
            try:
                tree.clear()
            finally:
                tree.blockSignals(False)

            if self._mode == "class":
                if self._info is not None:
                    self._info.setText("반별로 학생을 선택할 수 있습니다. (반 체크 = 반 전체 선택)")
                q = ""
                if self._search is not None:
                    q = (self._search.text() or "").strip().lower()
                students_by_id: Dict[str, Student] = {}
                if self.repo is not None:
                    for s in self.repo.list_all():
                        if s.id:
                            students_by_id[str(s.id)] = s
                if not self._classes:
                    empty = QTreeWidgetItem(tree)
                    empty.setText(0, "등록된 반이 없습니다.")
                    empty.setFont(0, _font(10, extra_bold=True))
                    empty.setFlags(empty.flags() & ~Qt.ItemIsSelectable)
                    self._sync_count()
                    return
                for c in self._classes:
                    student_ids = list(c.student_ids or [])
                    kids: List[Student] = []
                    for sid in student_ids:
                        st = students_by_id.get(str(sid))
                        if st is None and self.repo is not None:
                            st = self.repo.find_by_id(str(sid))
                        if st is not None:
                            if q and q not in ((st.name or "").lower()):
                                continue
                            kids.append(st)
                    label = f"{c.grade or ''} {c.name or ''}  ({len(kids)}명)".strip() or "(반)"
                    group = QTreeWidgetItem(tree)
                    group.setText(0, label)
                    group.setFont(0, _font(10, extra_bold=True))
                    group.setFlags(group.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsTristate)
                    for st in sorted(kids, key=lambda x: (x.name or "")):
                        if not st.id:
                            continue
                        sid = str(st.id)
                        child = QTreeWidgetItem(group)
                        child.setText(0, (st.name or "").strip() or "(이름 없음)")
                        child.setData(0, Qt.UserRole, sid)
                        child.setFont(0, _font(10, bold=True))
                        child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                        child.setCheckState(0, Qt.Checked if sid in self._selected_ids else Qt.Unchecked)
                    total = len(kids)
                    checked = sum(1 for st in kids if str(st.id) in self._selected_ids)
                    if checked <= 0:
                        group.setCheckState(0, Qt.Unchecked)
                    elif checked >= total:
                        group.setCheckState(0, Qt.Checked)
                    else:
                        group.setCheckState(0, Qt.PartiallyChecked)
                    try:
                        group.setExpanded(False)
                    except Exception:
                        pass
                self._sync_count()
                return

            if self._info is not None:
                self._info.setText("학년별로 학생을 선택할 수 있습니다. (학년 체크 = 학년 전체 선택)")

            students = [s for s in self._filtered_students() if (s.id and (s.grade or "").strip() and (s.name or "").strip())]
            if not students:
                empty = QTreeWidgetItem(tree)
                empty.setText(0, "표시할 학생이 없습니다.")
                empty.setFont(0, _font(10, extra_bold=True))
                empty.setFlags(empty.flags() & ~Qt.ItemIsSelectable)
                self._sync_count()
                return

            by_grade: Dict[str, List[Student]] = {}
            for s in students:
                g = (s.grade or "").strip()
                by_grade.setdefault(g, []).append(s)

            for g in sorted(by_grade.keys(), key=_grade_rank):
                group = QTreeWidgetItem(tree)
                group.setText(0, (f"{g}  ({len(by_grade[g])}명)" if g else f"({len(by_grade[g])}명)"))
                group.setFont(0, _font(10, extra_bold=True))
                group.setFlags(group.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsTristate)

                # children (id 없는 학생은 트리에 넣지 않음)
                kids = sorted(by_grade[g], key=lambda x: (x.name or ""))
                kids_with_id = [st for st in kids if st.id]
                for st in kids_with_id:
                    sid = str(st.id)
                    child = QTreeWidgetItem(group)
                    child.setText(0, (st.name or "").strip() or "(이름 없음)")
                    child.setData(0, Qt.UserRole, sid)
                    child.setFont(0, _font(10, bold=True))
                    child.setFlags(child.flags() | Qt.ItemIsUserCheckable)
                    child.setCheckState(0, Qt.Checked if sid in self._selected_ids else Qt.Unchecked)

                # group check state sync (트리에 넣은 학생만 기준)
                total = len(kids_with_id)
                checked = sum(1 for st in kids_with_id if str(st.id) in self._selected_ids)
                if checked <= 0:
                    group.setCheckState(0, Qt.Unchecked)
                elif checked >= total:
                    group.setCheckState(0, Qt.Checked)
                else:
                    group.setCheckState(0, Qt.PartiallyChecked)

                try:
                    group.setExpanded(False)
                except Exception:
                    pass

            self._sync_count()
        finally:
            self._building = False

    def _on_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        if self._building:
            return

        # 학생 item: Qt.UserRole에 id가 있음
        try:
            sid = item.data(0, Qt.UserRole)
        except Exception:
            sid = None

        # group item: children에 반영 (시그널 차단으로 clear/setCheckState 중 크래시 방지)
        if not sid:
            state = item.checkState(0)
            if state in (Qt.Checked, Qt.Unchecked):
                self._building = True
                tree = self._tree
                if tree is not None:
                    tree.blockSignals(True)
                try:
                    for i in range(item.childCount()):
                        ch = item.child(i)
                        if ch is None:
                            continue
                        try:
                            ch_id = ch.data(0, Qt.UserRole)
                        except Exception:
                            ch_id = None
                        ch_id = str(ch_id).strip() if ch_id is not None else ""
                        if ch_id and ch_id != "None":
                            if state == Qt.Checked:
                                self._selected_ids.add(ch_id)
                            else:
                                self._selected_ids.discard(ch_id)
                        ch.setCheckState(0, state)
                finally:
                    if tree is not None:
                        tree.blockSignals(False)
                    self._building = False
            self._sync_count()
            return

        sid = str(sid).strip() if sid is not None else ""
        if not sid or sid == "None":
            self._sync_count()
            return
        if item.checkState(0) == Qt.Checked:
            self._selected_ids.add(sid)
        else:
            self._selected_ids.discard(sid)
        self._sync_count()

    def _sync_count(self) -> None:
        if self._lbl_count is not None:
            self._lbl_count.setText(f"선택 {len(self._selected_ids)}명")

    def _on_ok(self) -> None:
        self.accept()


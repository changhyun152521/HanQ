"""
수업 탭 화면

요구사항:
- 메인 상단바/메인 사이드바는 건드리지 않고, '수업' 탭 콘텐츠 영역만 구현
- 측면 메뉴(탭 내부)에서 '학년' / '반'을 분리 선택
- 학년 선택 시 1번 스크린샷처럼: 학년 그룹 + 학생 리스트(출석 버튼) 형태로 로드
- 반 선택 시: 반별 학생 로드
- 반 정렬: 초4, 초5, 중1, 중2, 중3, 고1, 고2, 고3 (학년이 올라갈수록 뒤로)

백엔드 미연동 단계이므로, 화면 동작은 목데이터로 구성합니다.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont, QPalette
from PyQt5.QtWidgets import (
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QFrame,
    QPushButton,
    QLabel,
    QLineEdit,
    QTreeWidget,
    QTreeWidgetItem,
)

from core.models import Student
from database.repositories.student_repository import StudentRepository
from database.repositories.class_repository import ClassRepository
from ui.screens.student_page import StudentPage


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


CLASS_ORDER: List[str] = ["초1", "초2", "초3", "초4", "초5", "초6", "중1", "중2", "중3", "고1", "고2", "고3"]


@dataclass
class StudentItem:
    name: str
    grade: str  # "초4", "중3" 등
    id: Optional[str] = None


def _grade_rank(grade: str) -> int:
    """
    저학년이 위로 오도록 정렬 키를 반환합니다.
    - 초(0~99) < 중(100~199) < 고(200~299)
    - 숫자가 없거나 포맷이 이상하면 뒤로 보냅니다.
    """
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
    # 숫자 추출
    n = ""
    for ch in g:
        if ch.isdigit():
            n += ch
    try:
        return base + int(n or 99)
    except Exception:
        return base + 99


class ClassWorksheetScreen(QWidget):
    """수업 탭 메인 화면(탭 내부 사이드바 + 콘텐츠)."""

    def __init__(self, db_connection=None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db_connection = db_connection
        self.repo: Optional[StudentRepository] = None
        self._class_repo: Optional[ClassRepository] = None
        try:
            if self.db_connection is not None and getattr(self.db_connection, "is_connected", None) and self.db_connection.is_connected():
                self.repo = StudentRepository(self.db_connection)
                self._class_repo = ClassRepository(self.db_connection)
        except Exception:
            self.repo = None
            self._class_repo = None

        self._mode: str = "grade"  # "grade" | "class"
        self._students: List[StudentItem] = []
        self._selected_key: str = ""  # grade or class key

        self._tree: Optional[QTreeWidget] = None
        self._search: Optional[QLineEdit] = None
        self._right_title: Optional[QLabel] = None
        self._right_hint: Optional[QLabel] = None
        self._student_page: Optional[StudentPage] = None
        self._right_body = None
        self._right_body_lay = None

        self._build_ui()
        self._load_students_from_db()
        self._reload_sidebar()
        self._render_right_panel()

    def _load_students_from_db(self) -> None:
        """
        관리 > 학생관리에서 등록된 학생을 로드합니다.
        - MongoDB 연결 시: students 컬렉션에서 로드
        - 오프라인/미연결 시: 빈 목록(안내만 표시)
        """
        self._students = []
        if self.repo is None:
            return
        try:
            items = self.repo.list_all()
        except Exception:
            items = []
        for s in (items or []):
            if not s:
                continue
            name = (s.name or "").strip()
            grade = (s.grade or "").strip()
            if not name or not grade:
                continue
            self._students.append(StudentItem(name=name, grade=grade, id=(s.id or None)))

    def refresh_from_db(self) -> None:
        """DB에서 학생/반 목록을 다시 읽어 사이드바를 갱신. (관리에서 등록 후 수업 탭에서 바로 반영용)"""
        self._load_students_from_db()
        self._reload_sidebar()

    def _build_ui(self) -> None:
        self.setObjectName("ClassWorksheetRoot")
        self.setStyleSheet(
            """
            QWidget#ClassWorksheetRoot {
                background-color: #F8FAFC;
            }

            QFrame#InnerSidebar {
                background-color: #FFFFFF;
                border: 1px solid #E2E8F0;
                border-radius: 14px;
            }

            QFrame#InnerSidebar * {
                outline: none;
            }

            QFrame#SegmentWrap {
                background: #F1F5F9;
                border: 1px solid #F1F5F9;
                border-radius: 12px;
            }
            QPushButton#SegmentBtnLeft, QPushButton#SegmentBtnRight {
                border: none;
                border-radius: 12px;
                padding: 10px 0px;
                color: #334155;
                background: transparent;
                font-weight: 800;
            }
            QPushButton#SegmentBtnLeft:checked, QPushButton#SegmentBtnRight:checked {
                background: #2563EB;
                color: #FFFFFF;
                font-weight: 900;
            }

            QLineEdit#StudentSearch {
                background-color: #FFFFFF;
                border: 1px solid #475569; /* 더 진한 테두리 */
                border-radius: 12px;
                padding: 10px 12px;
                color: #0F172A; /* 더 진한 글씨 */
                font-weight: 800;
            }
            QLineEdit#StudentSearch:focus {
                border: 2px solid #2563EB;
            }
            QLineEdit#StudentSearch::placeholder {
                color: #64748B;
            }

            QTreeWidget#RosterTree {
                background: transparent;
                border: none;
            }
            /* 1. 아이템 스타일 (::branch는 건드리지 않아 Qt 기본 ▶/▼ 화살표가 그려지도록 함) */
            QTreeWidget#RosterTree::item {
                padding: 8px 5px;
                border: none;
                border-radius: 10px;
                color: #0F172A;
            }
            QTreeWidget#RosterTree::item:hover {
                background-color: #E8F0FE;
            }
            QTreeWidget#RosterTree::item:selected {
                background-color: #E8F0FE;
                color: #007BFF;
                font-weight: bold;
            }

            QFrame#RightCard {
                background-color: #FFFFFF;
                border: 1px solid #F1F5F9;
                border-radius: 14px;
            }
            /* 우측 본문 영역(학생 페이지 등): 전역 QWidget #F8FAFC 충돌 방지 → 흰색 고정 */
            QWidget#RightBody {
                background-color: #FFFFFF;
            }
            /* ✅ 전역 테마에서 QWidget 배경(#F8FAFC)이 QLabel에 적용되어 "연회색 바"처럼 보일 수 있어,
               우측 텍스트 영역은 항상 투명 배경을 강제합니다. */
            QFrame#RightCard QLabel {
                background: transparent;
            }
            QLabel#RightTitle {
                color: #0F172A;
            }
            QLabel#RightHint {
                color: #64748B;
            }
            """
        )

        outer = QHBoxLayout(self)
        outer.setContentsMargins(24, 16, 24, 24)
        outer.setSpacing(16)

        # 좌측: 탭 내부 사이드바
        sidebar = QFrame()
        sidebar.setObjectName("InnerSidebar")
        # ✅ 요청: 사이드바 너비를 "많이" 축소
        sidebar.setFixedWidth(200)

        s_lay = QVBoxLayout(sidebar)
        # ✅ 내부 여백/간격 축소(항목이 스크롤 없이 들어오게)
        s_lay.setContentsMargins(12, 12, 12, 12)
        s_lay.setSpacing(10)

        # 상단: 학년/반 세그먼트
        seg = QFrame()
        seg.setObjectName("SegmentWrap")
        seg_lay = QHBoxLayout(seg)
        seg_lay.setContentsMargins(2, 2, 2, 2)
        seg_lay.setSpacing(2)

        self.btn_grade = QPushButton("학년")
        self.btn_grade.setObjectName("SegmentBtnLeft")
        self.btn_grade.setCheckable(True)
        self.btn_grade.setChecked(True)
        self.btn_grade.setFocusPolicy(Qt.NoFocus)
        self.btn_grade.setCursor(Qt.PointingHandCursor)

        self.btn_class = QPushButton("반")
        self.btn_class.setObjectName("SegmentBtnRight")
        self.btn_class.setCheckable(True)
        self.btn_class.setChecked(False)
        self.btn_class.setFocusPolicy(Qt.NoFocus)
        self.btn_class.setCursor(Qt.PointingHandCursor)

        self.btn_grade.clicked.connect(lambda: self._set_mode("grade"))
        self.btn_class.clicked.connect(lambda: self._set_mode("class"))

        seg_lay.addWidget(self.btn_grade, 1)
        seg_lay.addWidget(self.btn_class, 1)
        s_lay.addWidget(seg)

        # 검색
        self._search = QLineEdit()
        self._search.setObjectName("StudentSearch")
        self._search.setPlaceholderText("학생 이름 검색")
        self._search.setClearButtonEnabled(True)
        self._search.setFixedHeight(40)
        self._search.textChanged.connect(self._reload_sidebar)
        s_lay.addWidget(self._search)

        # 리스트(스크롤)
        tree = QTreeWidget()
        self._tree = tree
        tree.setObjectName("RosterTree")
        tree.setHeaderHidden(True)
        tree.setRootIsDecorated(True)  # 확장/축소 화살표(▶/▼) 표시
        tree.setColumnCount(1)
        try:
            tree.setIndentation(20)  # 화살표와 텍스트 간격 확보
        except Exception:
            pass
        tree.itemClicked.connect(self._on_tree_item_clicked)
        tree.viewport().installEventFilter(self)
        # Qt 기본 브랜치 화살표(▶/▼)가 잘 보이도록 팔레트 색상 설정
        pal = tree.palette()
        dark = QColor(0x47, 0x56, 0x69)
        pal.setColor(QPalette.Text, dark)
        pal.setColor(QPalette.WindowText, dark)
        pal.setColor(QPalette.ButtonText, dark)
        tree.setPalette(pal)
        s_lay.addWidget(tree, 1)

        outer.addWidget(sidebar, 0)

        # 우측: 상세 패널(간단)
        right = QFrame()
        right.setObjectName("RightCard")
        r_lay = QVBoxLayout(right)
        r_lay.setContentsMargins(20, 18, 20, 18)
        r_lay.setSpacing(10)

        self._right_title = QLabel("수업")
        self._right_title.setObjectName("RightTitle")
        self._right_title.setFont(_font(14, extra_bold=True))
        r_lay.addWidget(self._right_title)

        self._right_hint = QLabel("좌측에서 학생을 선택하면 학생별 관리 페이지가 열립니다.")
        self._right_hint.setObjectName("RightHint")
        self._right_hint.setFont(_font(10, bold=True))
        self._right_hint.setWordWrap(True)
        r_lay.addWidget(self._right_hint)

        # 학생 선택 시 여기에 학생별 관리 페이지를 붙입니다.
        self._right_body = QWidget()
        self._right_body.setObjectName("RightBody")
        self._right_body_lay = QVBoxLayout(self._right_body)
        self._right_body_lay.setContentsMargins(0, 0, 0, 0)
        self._right_body_lay.setSpacing(0)
        r_lay.addWidget(self._right_body, 1)

        outer.addWidget(right, 1)

    def eventFilter(self, obj, event):
        """트리 viewport 이벤트 전달(클릭 영역 확장: 어디를 눌러도 펼침/접힘·선택 동작)."""
        return super().eventFilter(obj, event)

    def _set_mode(self, mode: str) -> None:
        m = "class" if (mode or "").strip() == "class" else "grade"
        self._mode = m
        self.btn_grade.setChecked(m == "grade")
        self.btn_class.setChecked(m == "class")
        self._selected_key = ""
        self._reload_sidebar()
        self._render_right_panel()

    def _filtered_students(self) -> List[StudentItem]:
        q = ""
        if self._search is not None:
            q = (self._search.text() or "").strip()
        if not q:
            return list(self._students)
        q2 = q.lower()
        return [s for s in self._students if q2 in (s.name or "").lower()]

    def _reload_sidebar(self) -> None:
        t = self._tree
        if t is None:
            return
        t.clear()

        # 반 탭: 반관리에서 등록한 반 로드 → 반별 학생 표시
        if self._mode == "class":
            if self._class_repo is None:
                info = QTreeWidgetItem(t)
                info.setText(0, "DB 연결 후 반 목록이 표시됩니다.")
                info.setFont(0, _font(10, extra_bold=True))
                return
            try:
                classes = self._class_repo.list_all()
            except Exception:
                classes = []
            id_to_student: Dict[str, StudentItem] = {str(s.id): s for s in self._students if s.id}
            q = (self._search.text() or "").strip().lower() if self._search else ""
            for c in sorted(classes, key=lambda x: (CLASS_ORDER.index(x.grade) if x.grade in CLASS_ORDER else 999, x.name or "")):
                n = len(c.student_ids or [])
                kids: List[StudentItem] = []
                for sid in (c.student_ids or []):
                    st = id_to_student.get(str(sid))
                    if st and (not q or q in (st.name or "").lower()):
                        kids.append(st)
                if q and not kids:
                    continue
                top = QTreeWidgetItem(t)
                top.setText(0, f"{c.name or ''}  ({len(kids)}명)")
                top.setData(0, Qt.UserRole, ("class", c.id, c.grade or "", c.name or ""))
                top.setFont(0, _font(10, extra_bold=True))
                for st in sorted(kids, key=lambda x: (x.name or "")):
                    it = QTreeWidgetItem(top)
                    it.setText(0, st.name)
                    it.setData(0, Qt.UserRole, ("student", st.grade or "", st.name or "", st.id or ""))
                    it.setFont(0, _font(10, bold=True))
                try:
                    top.setExpanded(False)
                except Exception:
                    pass
            if t.topLevelItemCount() == 0:
                empty = QTreeWidgetItem(t)
                empty.setText(0, "등록된 반이 없습니다")
                empty.setFont(0, _font(10, extra_bold=True))
            return

        students = self._filtered_students()
        if not students:
            empty = QTreeWidgetItem(t)
            empty.setText(0, "등록된 학생이 없습니다")
            empty.setFont(0, _font(10, extra_bold=True))
            return

        by_cls: Dict[str, List[StudentItem]] = {}
        for s in students:
            key = (s.grade or "").strip()
            if not key:
                continue
            by_cls.setdefault(key, []).append(s)

        # 표시 대상 키 정렬
        keys = sorted(by_cls.keys(), key=_grade_rank)

        for k in keys:
            top = QTreeWidgetItem(t)
            # ✅ 사이드바 폭 축소 대응: 카운트는 같은 라인에 붙여서 1컬럼 유지
            top.setText(0, f"{k}  ({len(by_cls.get(k, []))}명)")
            top.setData(0, Qt.UserRole, ("group", k))
            top.setFont(0, _font(10, extra_bold=True))

            for st in sorted(by_cls.get(k, []), key=lambda x: (x.name or "")):
                it = QTreeWidgetItem(top)
                it.setText(0, st.name)
                it.setData(0, Qt.UserRole, ("student", k, st.name, st.id))
                it.setFont(0, _font(10, bold=True))

            # ✅ 요청: 드롭다운을 미리 펼치지 않음(기본 접힘)
            try:
                top.setExpanded(False)
            except Exception:
                pass

        # 기본 선택은 비움(사용자가 클릭했을 때만 우측 패널 변경)
        self._selected_key = ""

    def _expand_group(self, key: str) -> None:
        t = self._tree
        if t is None:
            return
        for i in range(t.topLevelItemCount()):
            item = t.topLevelItem(i)
            if item is None:
                continue
            data = item.data(0, Qt.UserRole)
            if isinstance(data, tuple) and len(data) >= 2 and data[0] == "group" and str(data[1]) == str(key):
                try:
                    item.setExpanded(True)
                    t.setCurrentItem(item)
                except Exception:
                    pass
                break

    def _on_tree_item_clicked(self, item: QTreeWidgetItem, column: int) -> None:
        try:
            data = item.data(0, Qt.UserRole)
        except Exception:
            data = None

        if isinstance(data, tuple) and data:
            kind = data[0]
            if kind == "group":
                key = str(data[1])
                self._selected_key = key
                try:
                    item.setExpanded(not item.isExpanded())
                except Exception:
                    pass
                self._render_right_panel()
            elif kind == "class":
                self._selected_key = str(data[1] or "")
                try:
                    item.setExpanded(not item.isExpanded())
                except Exception:
                    pass
                self._render_right_panel()
            elif kind == "student":
                try:
                    grade = str(data[1] or "")
                    name = str(data[2] or "")
                    sid = str(data[3] or "")
                except Exception:
                    grade, name, sid = "", "", ""
                self._open_student_page(name=name, grade=grade, student_id=sid)

    def _open_student_page(self, *, name: str, grade: str, student_id: str) -> None:
        nm = (name or "").strip()
        gd = (grade or "").strip()
        sid = (student_id or "").strip()
        if not nm:
            return

        # 기존 페이지 제거 후 교체
        try:
            if self._student_page is not None:
                self._student_page.setParent(None)
        except Exception:
            pass

        self._student_page = StudentPage(self.db_connection, student_id=sid, student_name=nm, student_grade=gd)
        try:
            if self._right_body_lay is not None:
                self._right_body_lay.addWidget(self._student_page, 1)
        except Exception:
            pass

        # 타이틀/힌트는 숨김(학생 페이지가 대신 표시)
        try:
            if self._right_title is not None:
                self._right_title.hide()
            if self._right_hint is not None:
                self._right_hint.hide()
        except Exception:
            pass

    def _render_right_panel(self) -> None:
        if self._student_page is not None:
            return
        title = "수업"
        hint = "좌측에서 학생을 선택하면 학생별 관리 페이지가 열립니다."
        if self._right_title is not None:
            self._right_title.setText(title)
            self._right_title.show()
        if self._right_hint is not None:
            self._right_hint.setText(hint)
            self._right_hint.show()


"""
출처 선택 다이얼로그 모음

- Textbook 다중선택(단원 선택 상태에 맞는 교재만 표시)
- Exam(내신기출) 다중선택(연도/학년/학기/유형/학교명 필터, 선택 유지)
"""

from __future__ import annotations

from typing import Dict, List, Optional, Sequence, Set, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from database.sqlite_connection import SQLiteConnection
from database.repositories import ExamRepository, TextbookRepository
from core.models import Exam, Textbook
from services.worksheet import UnitKey


def _font(size: int, bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size))
    if bold:
        f.setBold(True)
    return f


class TextbookMultiSelectDialog(QDialog):
    """
    교재(Textbook) 다중 선택 다이얼로그

    - 단원 선택(UnitKey 리스트) 기준으로 subject/major/sub가 일치하는 Textbook만 노출
    - 필터(검색) 변경해도 선택한 ID는 유지
    """

    def __init__(
        self,
        db: SQLiteConnection,
        *,
        units: Sequence[UnitKey],
        preselected_ids: Optional[Sequence[str]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("교재 선택")
        self.setMinimumWidth(820)
        self.setMinimumHeight(540)

        self._db = db
        self._repo = TextbookRepository(db)
        self._units = [u.normalized() for u in units if u and u.is_valid()]
        self._unit_set = {(u.subject, u.major_unit, u.sub_unit) for u in self._units}

        self._all: List[Textbook] = []
        self._selected_ids: Set[str] = set(str(x) for x in (preselected_ids or []) if x)
        self._skip_next_selection_sync: bool = False

        self._init_ui()
        self._load()
        self._refresh()

    def selected_ids(self) -> List[str]:
        return list(self._selected_ids)

    def _init_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        title = QLabel("교재 선택 (단원 일치 데이터만 표시)")
        title.setFont(_font(12, bold=True))
        root.addWidget(title)

        top = QHBoxLayout()
        top.setSpacing(10)

        self.search = QLineEdit()
        self.search.setPlaceholderText("교재명 검색...")
        self.search.setClearButtonEnabled(True)
        self.search.setMinimumHeight(34)
        self.search.setFont(_font(10))
        self.search.textChanged.connect(self._refresh)
        top.addWidget(self.search, 1)

        self.lbl_count = QLabel("선택 0")
        self.lbl_count.setFont(_font(10, bold=True))
        top.addWidget(self.lbl_count)

        root.addLayout(top)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["선택", "교재명", "과목", "대단원", "소단원", "상태", "문항수"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(False)
        self.table.setColumnWidth(0, 56)
        self.table.setColumnWidth(1, 220)
        self.table.setColumnWidth(2, 140)
        self.table.setColumnWidth(3, 160)
        self.table.setColumnWidth(4, 160)
        self.table.setColumnWidth(5, 90)
        self.table.setColumnWidth(6, 70)
        self.table.setAlternatingRowColors(False)
        self.table.setCursor(Qt.PointingHandCursor)
        self.table.setStyleSheet(
            """
            QTableWidget {
                background-color: #FFFFFF;
                alternate-background-color: #FFFFFF;
            }
            QTableWidget::item {
                padding: 6px 4px;
                color: #222222;
                background-color: #FFFFFF;
                border: none;
                border-bottom: 1px solid #F0F0F0;
            }
            QTableWidget::item:hover { background-color: #f5f5f5; }
            QTableWidget::item:selected { background-color: #E8F2FF; }
            """
        )
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.table.selectionModel().selectionChanged.connect(self._on_selection_changed)
        root.addWidget(self.table, 1)

        btns = QHBoxLayout()
        btns.addStretch(1)

        btn_clear = QPushButton("전체 해제")
        btn_clear.setMinimumHeight(34)
        btn_clear.clicked.connect(self._clear_all)
        btns.addWidget(btn_clear)

        btn_cancel = QPushButton("취소")
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btns.addWidget(btn_cancel)

        btn_ok = QPushButton("확인")
        btn_ok.setMinimumHeight(34)
        btn_ok.setFont(_font(10, bold=True))
        btn_ok.clicked.connect(self.accept)
        btns.addWidget(btn_ok)

        root.addLayout(btns)

    def _load(self) -> None:
        self._all = self._repo.list_all()

    def _match_unit(self, tb: Textbook) -> bool:
        if not self._unit_set:
            return False
        s = (tb.subject or "").strip()
        m = (tb.major_unit or "").strip()
        sub = (tb.sub_unit or "").strip()
        return (s, m, sub) in self._unit_set

    def _refresh(self) -> None:
        q = (self.search.text() or "").strip().lower()

        rows: List[Textbook] = []
        for tb in self._all:
            if not tb or not tb.id:
                continue
            if not self._match_unit(tb):
                continue
            if q:
                if q not in (tb.name or "").lower():
                    continue
            rows.append(tb)

        # populate while blocking to avoid recursive itemChanged
        try:
            self.table.itemChanged.disconnect(self._on_item_changed)
        except Exception:
            pass

        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))
        for r, tb in enumerate(rows):
            tbid = str(tb.id)

            # checkbox via item check-state
            it = QTableWidgetItem("")
            it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if tbid in self._selected_ids else Qt.Unchecked)
            it.setData(Qt.UserRole, tbid)
            self.table.setItem(r, 0, it)

            self.table.setItem(r, 1, QTableWidgetItem(tb.name or ""))
            self.table.setItem(r, 2, QTableWidgetItem(tb.subject or ""))
            self.table.setItem(r, 3, QTableWidgetItem(tb.major_unit or ""))
            self.table.setItem(r, 4, QTableWidgetItem(tb.sub_unit or ""))

            status = "완료" if tb.is_parsed else "미완료"
            self.table.setItem(r, 5, QTableWidgetItem(status))
            self.table.setItem(r, 6, QTableWidgetItem(str(tb.problem_count or 0)))
        self.table.blockSignals(False)
        self.table.itemChanged.connect(self._on_item_changed)

        self._sync_count()

    def _on_cell_clicked(self, row: int, column: int) -> None:
        """행 클릭 시 체크박스 토글(단일 클릭만; 드래그는 _on_selection_changed에서 처리)"""
        self._skip_next_selection_sync = True
        it = self.table.item(row, 0)
        if it and (it.flags() & Qt.ItemIsUserCheckable):
            it.setCheckState(Qt.Unchecked if it.checkState() == Qt.Checked else Qt.Checked)

    def _on_selection_changed(self) -> None:
        """드래그 등으로 선택 영역이 바뀌면 선택된 행들의 체크박스를 모두 체크"""
        if self._skip_next_selection_sync:
            self._skip_next_selection_sync = False
            return
        rows = {idx.row() for idx in self.table.selectionModel().selectedIndexes()}
        for row in rows:
            it = self.table.item(row, 0)
            if it and (it.flags() & Qt.ItemIsUserCheckable):
                it.setCheckState(Qt.Checked)

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if not item or item.column() != 0:
            return
        tbid = item.data(Qt.UserRole)
        if not tbid:
            return
        if item.checkState() == Qt.Checked:
            self._selected_ids.add(str(tbid))
        else:
            self._selected_ids.discard(str(tbid))
        self._sync_count()

    def _sync_count(self) -> None:
        self.lbl_count.setText(f"선택 {len(self._selected_ids)}")

    def _clear_all(self) -> None:
        self._selected_ids.clear()
        self._refresh()


class ExamMultiSelectDialog(QDialog):
    """
    내신기출(Exam) 다중 선택 다이얼로그

    - 연도/학년/학기/유형/학교명(검색) 필터 제공
    - 필터를 바꿔도 선택된 Exam.id는 유지
    """

    def __init__(
        self,
        db: SQLiteConnection,
        *,
        preselected_ids: Optional[Sequence[str]] = None,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setWindowTitle("내신기출 선택")
        self.setMinimumWidth(920)
        self.setMinimumHeight(560)

        self._db = db
        self._repo = ExamRepository(db)
        self._all: List[Exam] = []
        self._selected_ids: Set[str] = set(str(x) for x in (preselected_ids or []) if x)
        self._skip_next_selection_sync: bool = False

        self._init_ui()
        self._load()
        self._populate_filters()
        self._refresh()

    def selected_ids(self) -> List[str]:
        return list(self._selected_ids)

    def _init_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(10)

        title = QLabel("내신기출 선택 (시험 행 다중선택, 필터 변경해도 선택 유지)")
        title.setFont(_font(12, bold=True))
        root.addWidget(title)

        filters = QHBoxLayout()
        filters.setSpacing(8)

        def _combo(label: str) -> Tuple[QLabel, QComboBox]:
            l = QLabel(label)
            l.setFont(_font(9, bold=True))
            c = QComboBox()
            c.setMinimumHeight(32)
            c.setFont(_font(9))
            c.addItem("전체")
            c.currentTextChanged.connect(self._refresh)
            return l, c

        l, self.year = _combo("연도")
        filters.addWidget(l)
        filters.addWidget(self.year)

        l, self.grade = _combo("학년")
        filters.addWidget(l)
        filters.addWidget(self.grade)

        l, self.semester = _combo("학기")
        filters.addWidget(l)
        filters.addWidget(self.semester)

        l, self.exam_type = _combo("유형")
        filters.addWidget(l)
        filters.addWidget(self.exam_type)

        self.school_search = QLineEdit()
        self.school_search.setPlaceholderText("학교명 검색...")
        self.school_search.setClearButtonEnabled(True)
        self.school_search.setMinimumHeight(32)
        self.school_search.setFont(_font(9))
        self.school_search.textChanged.connect(self._refresh)
        filters.addWidget(self.school_search, 1)

        self.lbl_count = QLabel("선택 0")
        self.lbl_count.setFont(_font(10, bold=True))
        filters.addWidget(self.lbl_count)

        root.addLayout(filters)

        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["선택", "연도", "학년", "학기", "유형", "학교", "상태", "문항수"])
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.verticalHeader().setVisible(False)
        self.table.setAlternatingRowColors(False)
        self.table.setColumnWidth(0, 56)
        self.table.setColumnWidth(1, 70)
        self.table.setColumnWidth(2, 70)
        self.table.setColumnWidth(3, 70)
        self.table.setColumnWidth(4, 120)
        self.table.setColumnWidth(5, 260)
        self.table.setColumnWidth(6, 90)
        self.table.setColumnWidth(7, 70)
        self.table.setCursor(Qt.PointingHandCursor)
        self.table.setStyleSheet(
            """
            QTableWidget {
                background-color: #FFFFFF;
                alternate-background-color: #FFFFFF;
            }
            QTableWidget::item {
                padding: 6px 4px;
                color: #222222;
                background-color: #FFFFFF;
                border: none;
                border-bottom: 1px solid #F0F0F0;
            }
            QTableWidget::item:hover { background-color: #f5f5f5; }
            QTableWidget::item:selected { background-color: #E8F2FF; }
            """
        )
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.table.selectionModel().selectionChanged.connect(self._on_selection_changed)
        root.addWidget(self.table, 1)

        btns = QHBoxLayout()
        btns.addStretch(1)

        btn_clear = QPushButton("전체 해제")
        btn_clear.setMinimumHeight(34)
        btn_clear.clicked.connect(self._clear_all)
        btns.addWidget(btn_clear)

        btn_cancel = QPushButton("취소")
        btn_cancel.setMinimumHeight(34)
        btn_cancel.clicked.connect(self.reject)
        btns.addWidget(btn_cancel)

        btn_ok = QPushButton("확인")
        btn_ok.setMinimumHeight(34)
        btn_ok.setFont(_font(10, bold=True))
        btn_ok.clicked.connect(self.accept)
        btns.addWidget(btn_ok)

        root.addLayout(btns)

    def _load(self) -> None:
        self._all = self._repo.list_all()

    def _populate_filters(self) -> None:
        def uniq(getter):
            vals = sorted({(getter(x) or "").strip() for x in self._all if x} - {""})
            return vals

        for combo, vals in (
            (self.year, uniq(lambda e: e.year)),
            (self.grade, uniq(lambda e: e.grade)),
            (self.semester, uniq(lambda e: e.semester)),
            (self.exam_type, uniq(lambda e: e.exam_type)),
        ):
            current = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            combo.addItem("전체")
            for v in vals:
                combo.addItem(v)
            if current and current != "전체" and current in vals:
                combo.setCurrentText(current)
            combo.blockSignals(False)

    def _match(self, e: Exam) -> bool:
        if not e or not e.id:
            return False
        y = self.year.currentText()
        g = self.grade.currentText()
        s = self.semester.currentText()
        t = self.exam_type.currentText()
        q = (self.school_search.text() or "").strip().lower()

        if y != "전체" and (e.year or "").strip() != y:
            return False
        if g != "전체" and (e.grade or "").strip() != g:
            return False
        if s != "전체" and (e.semester or "").strip() != s:
            return False
        if t != "전체" and (e.exam_type or "").strip() != t:
            return False
        if q and q not in (e.school_name or "").lower():
            return False
        return True

    def _refresh(self) -> None:
        rows = [e for e in self._all if self._match(e)]
        try:
            self.table.itemChanged.disconnect(self._on_item_changed)
        except Exception:
            pass

        self.table.blockSignals(True)
        self.table.setRowCount(len(rows))

        for r, e in enumerate(rows):
            eid = str(e.id)
            it = QTableWidgetItem("")
            it.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if eid in self._selected_ids else Qt.Unchecked)
            it.setData(Qt.UserRole, eid)
            self.table.setItem(r, 0, it)

            self.table.setItem(r, 1, QTableWidgetItem(e.year or ""))
            self.table.setItem(r, 2, QTableWidgetItem(e.grade or ""))
            self.table.setItem(r, 3, QTableWidgetItem(e.semester or ""))
            self.table.setItem(r, 4, QTableWidgetItem(e.exam_type or ""))
            self.table.setItem(r, 5, QTableWidgetItem(e.school_name or ""))

            status = "완료" if e.is_parsed else "미완료"
            self.table.setItem(r, 6, QTableWidgetItem(status))
            self.table.setItem(r, 7, QTableWidgetItem(str(e.problem_count or 0)))

        self.table.blockSignals(False)
        self.table.itemChanged.connect(self._on_item_changed)
        self._sync_count()

    def _on_cell_clicked(self, row: int, column: int) -> None:
        """행 클릭 시 체크박스 토글(단일 클릭만; 드래그는 _on_selection_changed에서 처리)"""
        self._skip_next_selection_sync = True
        it = self.table.item(row, 0)
        if it and (it.flags() & Qt.ItemIsUserCheckable):
            it.setCheckState(Qt.Unchecked if it.checkState() == Qt.Checked else Qt.Checked)

    def _on_selection_changed(self) -> None:
        """드래그 등으로 선택 영역이 바뀌면 선택된 행들의 체크박스를 모두 체크"""
        if self._skip_next_selection_sync:
            self._skip_next_selection_sync = False
            return
        rows = {idx.row() for idx in self.table.selectionModel().selectedIndexes()}
        for row in rows:
            it = self.table.item(row, 0)
            if it and (it.flags() & Qt.ItemIsUserCheckable):
                it.setCheckState(Qt.Checked)

    def _on_item_changed(self, item: QTableWidgetItem) -> None:
        if not item or item.column() != 0:
            return
        eid = item.data(Qt.UserRole)
        if not eid:
            return
        if item.checkState() == Qt.Checked:
            self._selected_ids.add(str(eid))
        else:
            self._selected_ids.discard(str(eid))
        self._sync_count()

    def _sync_count(self) -> None:
        self.lbl_count.setText(f"선택 {len(self._selected_ids)}")

    def _clear_all(self) -> None:
        self._selected_ids.clear()
        self._refresh()


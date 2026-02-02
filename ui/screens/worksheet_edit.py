"""
학습지 문항 미리보기/편집 화면 (좌: 미리보기 / 우: 문항 목록)

요구사항:
- 좌측: 선택된 1문항 미리보기(텍스트) + 원본 HWP 보기
- 우측: 문항 리스트(문항번호 + 단원/난이도/출처만), 드래그로 순서 변경, 삭제
- 문항번호는 "최종 생성 시점"에 확정되어 mapping(no -> problem_id)로 데이터화(후속 오답노트용)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional
import os
from datetime import datetime

from PyQt5.QtCore import Qt, QSize, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QPainter, QPen, QTextCursor, QTextBlockFormat
from PyQt5.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
    QGraphicsDropShadowEffect,
    QLineEdit,
    QCheckBox,
)

from core.models import SourceType
from database.sqlite_connection import SQLiteConnection
from database.repositories import ExamRepository, TextbookRepository
from services.problem import ProblemService
from services.worksheet.hwp_composer import WorksheetComposeError, WorksheetHwpComposer
from utils.hwp_restore import HWPRestore


def _font(size: int, bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("NanumGothic")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size))
    if bold:
        f.setBold(True)
    return f


@dataclass
class WorksheetDraft:
    title: str
    creator: str
    grade: str
    type_text: str
    option_unit_tag: bool
    option_source_tag: bool
    option_difficulty_tag: bool
    requested_total: int
    actual_total: int
    warnings: List[str]
    problem_ids: List[str]


class _DraggableList(QListWidget):
    reordered = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QListWidget.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QListWidget.SingleSelection)
        self.setSpacing(8)

        # drag insertion indicator (custom blue line)
        self._indicator_y: Optional[int] = None
        try:
            self.model().rowsMoved.connect(self.reordered)  # type: ignore[attr-defined]
        except Exception:
            pass

    def dropEvent(self, event):  # noqa: N802
        super().dropEvent(event)
        self._indicator_y = None
        self.viewport().update()
        self.reordered.emit()

    def dragMoveEvent(self, event):  # noqa: N802
        super().dragMoveEvent(event)
        try:
            idx = self.indexAt(event.pos())
            if idx.isValid():
                rect = self.visualRect(idx)
                y = rect.top() if event.pos().y() < rect.center().y() else rect.bottom()
                self._indicator_y = int(y)
            else:
                # empty space: line at bottom
                self._indicator_y = int(self.viewport().height() - 2)
        except Exception:
            self._indicator_y = None
        self.viewport().update()

    def dragLeaveEvent(self, event):  # noqa: N802
        self._indicator_y = None
        self.viewport().update()
        super().dragLeaveEvent(event)

    def paintEvent(self, event):  # noqa: N802
        super().paintEvent(event)
        if self._indicator_y is None:
            return
        try:
            p = QPainter(self.viewport())
            p.setRenderHint(QPainter.Antialiasing, True)
            pen = QPen(QColor("#2563EB"))
            pen.setWidth(2)
            p.setPen(pen)
            x1 = 10
            x2 = max(x1 + 1, self.viewport().width() - 10)
            y = int(self._indicator_y)
            p.drawLine(x1, y, x2, y)
            p.end()
        except Exception:
            pass


class _ProblemListRow(QFrame):
    remove_requested = pyqtSignal(str)
    selected_requested = pyqtSignal(str)

    def __init__(
        self,
        *,
        problem_id: str,
        number: int,
        unit_text: str,
        difficulty: str,
        source_text: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.problem_id = problem_id
        self.setObjectName("Row")
        self.setCursor(Qt.PointingHandCursor)

        self._is_selected = False
        self._is_hovered = False

        root = QHBoxLayout(self)
        root.setContentsMargins(12, 8, 12, 8)
        root.setSpacing(10)

        self.lbl_no = QLabel(str(int(number)))
        self.lbl_no.setFont(_font(10, bold=True))
        self.lbl_no.setAlignment(Qt.AlignCenter)
        self.lbl_no.setFixedSize(24, 24)
        self.lbl_no.setStyleSheet(
            """
            QLabel {
                background-color: transparent;
                color: #2563EB;
                font-weight: 800;
            }
            """
        )
        root.addWidget(self.lbl_no, alignment=Qt.AlignVCenter)

        meta_col = QVBoxLayout()
        meta_col.setSpacing(4)

        line1 = QLabel((unit_text or "").strip())
        line1.setFont(_font(10, bold=True))
        line1.setStyleSheet("background: transparent; color: #222222;")
        line1.setWordWrap(True)

        line2_txt = " / ".join([x for x in [difficulty, source_text] if x])
        line2 = QLabel(line2_txt)
        line2.setFont(_font(9))
        line2.setStyleSheet("background: transparent; color: #222222;")
        line2.setWordWrap(True)

        meta_col.addWidget(line1)
        if line2_txt:
            meta_col.addWidget(line2)

        root.addLayout(meta_col, 1)

        self.btn_remove = QPushButton("×")
        self.btn_remove.setObjectName("RemoveBtn")
        self.btn_remove.setFixedSize(24, 24)
        self.btn_remove.setFont(_font(12, bold=True))
        self.btn_remove.setCursor(Qt.PointingHandCursor)
        self.btn_remove.clicked.connect(lambda: self.remove_requested.emit(self.problem_id))
        root.addWidget(self.btn_remove, alignment=Qt.AlignVCenter)

        self._apply_styles()

    def mousePressEvent(self, event):  # noqa: N802
        try:
            self.selected_requested.emit(self.problem_id)
        except Exception:
            pass
        super().mousePressEvent(event)

    def enterEvent(self, event):  # noqa: N802
        self._is_hovered = True
        self._apply_styles()
        super().enterEvent(event)

    def leaveEvent(self, event):  # noqa: N802
        self._is_hovered = False
        self._apply_styles()
        super().leaveEvent(event)

    def set_number(self, n: int) -> None:
        self.lbl_no.setText(str(int(n)))

    def set_selected(self, selected: bool) -> None:
        selected = bool(selected)
        if self._is_selected == selected:
            return
        self._is_selected = selected
        self._apply_styles()

    def _apply_styles(self) -> None:
        if self._is_selected:
            border_bottom = "#E8F2FF"
            bg = "transparent"
        elif self._is_hovered:
            border_bottom = "#F0F0F0"
            bg = "#FAFAFA"
        else:
            border_bottom = "#F0F0F0"
            bg = "transparent"

        self.setStyleSheet(
            f"""
            QFrame#Row {{
                background: {bg};
                border: none;
                border-bottom: 1px solid {border_bottom};
                border-radius: 0;
            }}
            QPushButton#RemoveBtn {{
                background-color: transparent;
                color: #FF4D4F;
                border: 1px solid #FF4D4F;
                border-radius: 4px;
                font-weight: bold;
                padding: 2px 6px;
            }}
            QPushButton#RemoveBtn:hover {{
                background-color: #FFF1F0;
                color: #FF4D4F;
            }}
            """
        )

class WorksheetEditScreen(QWidget):
    back_requested = pyqtSignal()
    # {'output_path': str, 'numbered': [{'no': int, 'problem_id': str}, ...]}
    finalized = pyqtSignal(dict)

    def __init__(self, db_connection: SQLiteConnection, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db = db_connection
        self.problem_service = ProblemService(db_connection)
        self.textbook_repo = TextbookRepository(db_connection)
        self.exam_repo = ExamRepository(db_connection)
        self.hwp_restore = HWPRestore(db_connection)

        self._draft: Optional[WorksheetDraft] = None
        self._unit: Dict[str, str] = {}
        self._diff: Dict[str, str] = {}
        self._source: Dict[str, str] = {}
        self._preview_text: Dict[str, str] = {}
        self._detail_cache: Dict[str, dict] = {}

        self._init_ui()

    def _init_ui(self) -> None:
        self.setObjectName("WorksheetEditScreen")
        self.setStyleSheet(
            """
            QWidget#WorksheetEditScreen QFrame#FinalConfig {
                background-color: #FFFFFF;
                border: none;
                border-radius: 12px;
            }
            QWidget#WorksheetEditScreen QFrame#ListPanel {
                background-color: transparent;
                border: none;
            }
            QWidget#WorksheetEditScreen QFrame#PreviewWrap,
            QWidget#WorksheetEditScreen QFrame#PreviewCanvas {
                background-color: #FFFFFF;
                border: none;
            }
            QWidget#WorksheetEditScreen QLabel#FinalConfigTitle {
                color: #222222;
                background: transparent;
            }
            QWidget#WorksheetEditScreen QLineEdit#FinalInput {
                background: transparent;
                border: none;
                border-bottom: 1px solid #E0E0E0;
                padding: 6px 0;
                color: #222222;
            }
            QWidget#WorksheetEditScreen QLineEdit#FinalInput:focus {
                border-bottom: 2px solid #2563EB;
            }
            QWidget#WorksheetEditScreen QCheckBox#FinalOpt {
                color: #222222;
                background: transparent;
            }
            QWidget#WorksheetEditScreen QCheckBox#FinalOpt::indicator {
                width: 18px;
                height: 18px;
                border: 2px solid #94A3B8;
                border-radius: 4px;
                background-color: #FFFFFF;
            }
            QWidget#WorksheetEditScreen QCheckBox#FinalOpt::indicator:hover {
                border-color: #64748B;
                background-color: #F8FAFC;
            }
            QWidget#WorksheetEditScreen QCheckBox#FinalOpt::indicator:checked {
                border-color: #2563EB;
                background-color: #2563EB;
            }
            QWidget#WorksheetEditScreen QListWidget {
                background-color: transparent;
                border: none;
                outline: none;
            }
            """
        )
        root = QVBoxLayout(self)
        root.setContentsMargins(28, 18, 28, 24)
        root.setSpacing(14)

        # 상단 헤더
        header = QHBoxLayout()
        header.setSpacing(10)

        self.lbl_title = QLabel("문항 미리보기 / 편집")
        self.lbl_title.setFont(_font(14, bold=True))
        self.lbl_title.setStyleSheet("background: transparent; color: #222222;")
        header.addWidget(self.lbl_title)

        self.lbl_count = QLabel("(요청 0 → 선택 0)")
        self.lbl_count.setFont(_font(10))
        self.lbl_count.setStyleSheet("background: transparent; color: #222222;")
        header.addWidget(self.lbl_count, alignment=Qt.AlignVCenter)

        header.addStretch(1)

        root.addLayout(header)

        # 중앙: 좌(미리보기) / 우(문항 목록)
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setStyleSheet("QSplitter::handle{background-color:#E2E8F0; width:2px;}")

        # 1) 좌측: 미리보기(약 35%)
        left = QFrame()
        left.setObjectName("PreviewWrap")
        l = QVBoxLayout(left)
        l.setContentsMargins(0, 0, 0, 0)
        l.setSpacing(0)

        self.canvas = QFrame()
        self.canvas.setObjectName("PreviewCanvas")
        self._apply_canvas_shadow(self.canvas)
        cl = QVBoxLayout(self.canvas)
        cl.setContentsMargins(18, 16, 18, 14)
        cl.setSpacing(10)

        self.preview_meta = QLabel("")
        self.preview_meta.setFont(_font(9))
        self.preview_meta.setStyleSheet("background: transparent; color: #222222;")
        self.preview_meta.setWordWrap(True)
        cl.addWidget(self.preview_meta)

        self.preview_box = QPlainTextEdit()
        self.preview_box.setReadOnly(True)
        self.preview_box.setFont(_font(9))
        self.preview_box.setStyleSheet(
            """
            QPlainTextEdit {
                background: transparent;
                border: none;
                padding: 0px;
                color: #222222;
            }
            """
        )
        cl.addWidget(self.preview_box, 1)

        lbtn = QHBoxLayout()
        lbtn.setContentsMargins(0, 0, 0, 0)
        lbtn.addStretch(1)
        self.btn_open_hwp = QPushButton("원본 HWP 보기")
        self.btn_open_hwp.setObjectName("OpenHwpTextBtn")
        self.btn_open_hwp.setCursor(Qt.PointingHandCursor)
        self.btn_open_hwp.setFont(_font(9, bold=True))
        self.btn_open_hwp.setFlat(True)
        self.btn_open_hwp.clicked.connect(self._on_open_original_clicked)
        lbtn.addWidget(self.btn_open_hwp, alignment=Qt.AlignRight)
        cl.addLayout(lbtn)

        l.addWidget(self.canvas, 1)

        # 2) 중앙: 문항 목록(약 35%) — 드래그 작업 공간
        mid = QFrame()
        mid.setObjectName("ListPanel")
        m = QVBoxLayout(mid)
        m.setContentsMargins(14, 14, 14, 14)
        m.setSpacing(10)

        self.list_title = QLabel("문항 목록 (드래그로 순서 변경)")
        self.list_title.setFont(_font(12, bold=True))
        self.list_title.setStyleSheet("background: transparent; color: #222222;")
        m.addWidget(self.list_title)

        self.list = _DraggableList()
        self.list.reordered.connect(self._on_reordered)
        self.list.currentItemChanged.connect(self._on_current_item_changed)
        m.addWidget(self.list, 1)

        # 3) 우측: 문서 설정(약 30%) — 시원하게
        cfg = QFrame()
        cfg.setObjectName("ConfigPanel")
        c = QVBoxLayout(cfg)
        c.setContentsMargins(14, 14, 14, 14)
        c.setSpacing(12)

        self.final_config = self._build_final_config_widget()
        c.addWidget(self.final_config, 0)

        c.addStretch(1)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch(1)

        self.btn_back = QPushButton("이전")
        self.btn_back.setObjectName("BackBtn")
        self.btn_back.setFont(_font(10, bold=True))
        self.btn_back.setFixedHeight(42)
        self.btn_back.clicked.connect(self.back_requested.emit)
        btn_row.addWidget(self.btn_back)

        self.btn_finalize = QPushButton("최종 생성(HWP)")
        self.btn_finalize.setObjectName("FinalizeBtn")
        self.btn_finalize.setFont(_font(10, bold=True))
        self.btn_finalize.setFixedHeight(42)
        self.btn_finalize.clicked.connect(self._on_finalize_clicked)
        btn_row.addWidget(self.btn_finalize)

        c.addLayout(btn_row)

        splitter.addWidget(left)
        splitter.addWidget(mid)
        splitter.addWidget(cfg)
        # 35% / 35% / 30% 비율감(7:7:6)
        splitter.setStretchFactor(0, 7)
        splitter.setStretchFactor(1, 7)
        splitter.setStretchFactor(2, 6)
        root.addWidget(splitter, 1)

        # 스타일은 앱 전역 테마(`ui/theme.py`)에서 관리합니다.

    def load_draft(self, draft: WorksheetDraft) -> None:
        self._draft = draft
        self._build_cache(draft.problem_ids)
        self._render_list(draft.problem_ids)
        self._sync_header()
        self._sync_final_config_from_draft()
        if self.list.count() > 0:
            self.list.setCurrentRow(0)

    def current_problem_ids(self) -> List[str]:
        ids: List[str] = []
        for i in range(self.list.count()):
            it = self.list.item(i)
            pid = it.data(Qt.UserRole)
            if pid:
                ids.append(str(pid))
        return ids

    def _sync_header(self) -> None:
        d = self._draft
        if not d:
            self.lbl_title.setText("문항 미리보기 / 편집")
            self.lbl_count.setText("(요청 0 → 선택 0)")
            return
        self.lbl_title.setText("문항 미리보기 / 편집")
        self.lbl_count.setText(f"(요청 {d.requested_total} → 선택 {self.list.count()})")

    def _build_final_config_widget(self) -> QFrame:
        card = QFrame()
        card.setObjectName("FinalConfig")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(14)

        title = QLabel("문서 설정")
        title.setObjectName("FinalConfigTitle")
        title.setFont(_font(11, bold=True))
        lay.addWidget(title)
        lay.addSpacing(4)

        # 학습지명
        lbl_t = QLabel("학습지명")
        lbl_t.setFont(_font(9, bold=True))
        lbl_t.setStyleSheet("background: transparent; color: #222222;")
        lay.addWidget(lbl_t)
        lay.addSpacing(4)
        self.title_input = QLineEdit()
        self.title_input.setObjectName("FinalInput")
        self.title_input.setPlaceholderText("예: 1학기 중간고사 대비 모의고사")
        self.title_input.setFont(_font(12, bold=False))
        self.title_input.setFixedHeight(40)
        lay.addWidget(self.title_input)

        line1 = QFrame()
        line1.setObjectName("DividerLine")
        line1.setFixedHeight(1)
        line1.setStyleSheet("background-color: #F0F0F0; border: none;")
        lay.addWidget(line1)
        lay.addSpacing(8)

        lbl_a = QLabel("출제자")
        lbl_a.setFont(_font(9, bold=True))
        lbl_a.setStyleSheet("background: transparent; color: #222222;")
        lay.addWidget(lbl_a)
        lay.addSpacing(4)
        self.author_input = QLineEdit()
        self.author_input.setObjectName("FinalInput")
        self.author_input.setPlaceholderText("예: 이창현T")
        self.author_input.setFont(_font(12, bold=False))
        self.author_input.setFixedHeight(40)
        lay.addWidget(self.author_input)

        line2 = QFrame()
        line2.setObjectName("DividerLine")
        line2.setFixedHeight(1)
        line2.setStyleSheet("background-color: #F0F0F0; border: none;")
        lay.addWidget(line2)
        lay.addSpacing(8)

        lbl_o = QLabel("표시 옵션")
        lbl_o.setFont(_font(9, bold=True))
        lbl_o.setStyleSheet("background: transparent; color: #222222;")
        lay.addWidget(lbl_o)
        lay.addSpacing(4)

        # 표시 옵션(세로) — 클릭 영역 확보
        opt_col = QVBoxLayout()
        opt_col.setSpacing(10)

        self.check_unit = QCheckBox("단원명")
        self.check_unit.setObjectName("FinalOpt")
        self.check_source = QCheckBox("출처")
        self.check_source.setObjectName("FinalOpt")
        self.check_level = QCheckBox("난이도")
        self.check_level.setObjectName("FinalOpt")
        for cb in (self.check_unit, self.check_source, self.check_level):
            cb.setFont(_font(11, bold=True))
            cb.setContentsMargins(2, 6, 2, 6)
            opt_col.addWidget(cb)
        lay.addLayout(opt_col)

        return card

    def _sync_final_config_from_draft(self) -> None:
        d = self._draft
        if not d:
            return
        try:
            self.title_input.setText(d.title or "")
            self.author_input.setText(d.creator or "")
            self.check_unit.setChecked(bool(d.option_unit_tag))
            self.check_source.setChecked(bool(d.option_source_tag))
            self.check_level.setChecked(bool(d.option_difficulty_tag))
        except Exception:
            pass

    def _apply_final_config_to_draft(self) -> None:
        """Step2 입력값을 draft에 반영 (데이터 구조 변경 없음)."""
        d = self._draft
        if not d:
            return
        try:
            d.title = (self.title_input.text() or "").strip()
            d.creator = (self.author_input.text() or "").strip()
            d.option_unit_tag = bool(self.check_unit.isChecked())
            d.option_source_tag = bool(self.check_source.isChecked())
            d.option_difficulty_tag = bool(self.check_level.isChecked())
        except Exception:
            pass

    def _apply_canvas_shadow(self, w: QFrame) -> None:
        # 아주 연한 그림자(종이 질감)
        try:
            shadow = QGraphicsDropShadowEffect(w)
            shadow.setBlurRadius(25)
            shadow.setXOffset(0)
            shadow.setYOffset(12)
            shadow.setColor(QColor(0, 0, 0, 26))  # opacity ~0.1
            w.setGraphicsEffect(shadow)
        except Exception:
            pass

    def _apply_preview_line_spacing(self) -> None:
        """미리보기 텍스트 줄 간격 1.6배 적용."""
        try:
            doc = self.preview_box.document()
            cur = QTextCursor(doc)
            cur.select(QTextCursor.Document)
            fmt = QTextBlockFormat()
            fmt.setLineHeight(160, QTextBlockFormat.ProportionalHeight)
            cur.mergeBlockFormat(fmt)
        except Exception:
            pass

    def _build_cache(self, problem_ids: List[str]) -> None:
        tb_map = {str(t.id): t for t in self.textbook_repo.list_all() if t and t.id}
        ex_map = {str(e.id): e for e in self.exam_repo.list_all() if e and e.id}

        self._unit.clear()
        self._diff.clear()
        self._source.clear()
        self._preview_text.clear()
        self._detail_cache.clear()

        for pid in problem_ids:
            detail = self.problem_service.get_problem_detail(pid)
            if not detail:
                self._unit[pid] = ""
                self._diff[pid] = ""
                self._source[pid] = ""
                self._preview_text[pid] = ""
                continue

            self._detail_cache[pid] = detail

            tags = detail.get("tags") or []
            t0 = tags[0] if tags else {}
            subject = (t0.get("subject") or "").strip()
            major = (t0.get("major_unit") or "").strip()
            sub = (t0.get("sub_unit") or "").strip()
            diff = (t0.get("difficulty") or "").strip()

            self._unit[pid] = f"{subject} / {major} > {sub}".strip()
            self._diff[pid] = diff

            st = detail.get("source_type")
            sid = detail.get("source_id")
            if st == SourceType.TEXTBOOK.value and sid:
                tb = tb_map.get(str(sid))
                self._source[pid] = tb.name if tb else ""
            elif st == SourceType.EXAM.value and sid:
                ex = ex_map.get(str(sid))
                self._source[pid] = f"{ex.school_name} {ex.grade} {ex.semester} {ex.exam_type} ({ex.year})" if ex else ""
            else:
                self._source[pid] = ""

            self._preview_text[pid] = (detail.get("content_text") or "").strip()

    def _render_list(self, problem_ids: List[str]) -> None:
        self.list.clear()
        for idx, pid in enumerate(problem_ids, start=1):
            item = QListWidgetItem()
            item.setData(Qt.UserRole, pid)
            item.setSizeHint(QSize(0, 62))
            item.setFlags(item.flags() | Qt.ItemIsDragEnabled | Qt.ItemIsEnabled | Qt.ItemIsSelectable)

            row = _ProblemListRow(
                problem_id=pid,
                number=idx,
                unit_text=self._unit.get(pid, ""),
                difficulty=self._diff.get(pid, ""),
                source_text=self._source.get(pid, ""),
            )
            row.remove_requested.connect(self._remove_problem)
            row.selected_requested.connect(self._set_current_by_id)

            self.list.addItem(item)
            self.list.setItemWidget(item, row)
        self._sync_list_row_selected_state()

    def _renumber(self) -> None:
        for i in range(self.list.count()):
            it = self.list.item(i)
            w = self.list.itemWidget(it)
            if isinstance(w, _ProblemListRow):
                w.set_number(i + 1)

    def _set_current_by_id(self, problem_id: str) -> None:
        for i in range(self.list.count()):
            it = self.list.item(i)
            if str(it.data(Qt.UserRole)) == str(problem_id):
                self.list.setCurrentRow(i)
                break

    def _remove_problem(self, problem_id: str) -> None:
        for i in range(self.list.count()):
            it = self.list.item(i)
            if str(it.data(Qt.UserRole)) == str(problem_id):
                self.list.takeItem(i)
                break
        self._renumber()
        self._sync_header()
        self._sync_list_row_selected_state()
        if self.list.count() == 0:
            self._show_preview(None)

    def _on_reordered(self) -> None:
        cur = self._current_problem_id()
        self._renumber()
        self._sync_header()
        # reorder 후에도 선택/하이라이트 유지
        if cur:
            self._set_current_by_id(cur)
        self._sync_list_row_selected_state()

    def _on_current_item_changed(self, current: Optional[QListWidgetItem], _prev: Optional[QListWidgetItem]) -> None:
        pid = str(current.data(Qt.UserRole)) if current is not None else None
        self._show_preview(pid)
        self._sync_list_row_selected_state()

    def _sync_list_row_selected_state(self) -> None:
        cur = self._current_problem_id()
        for i in range(self.list.count()):
            it = self.list.item(i)
            w = self.list.itemWidget(it)
            if isinstance(w, _ProblemListRow):
                w.set_selected(str(it.data(Qt.UserRole)) == str(cur) if cur else False)

    def _show_preview(self, problem_id: Optional[str]) -> None:
        if not problem_id:
            self.preview_meta.setText("")
            self.preview_box.setPlainText("")
            self.btn_open_hwp.setEnabled(False)
            return

        unit = self._unit.get(problem_id, "")
        diff = self._diff.get(problem_id, "")
        src = self._source.get(problem_id, "")
        self.preview_meta.setText(" | ".join([x for x in [unit, diff, src] if x]))
        self.preview_box.setPlainText(self._preview_text.get(problem_id, "") or "")
        self._apply_preview_line_spacing()
        self.btn_open_hwp.setEnabled(True)

    def _current_problem_id(self) -> Optional[str]:
        it = self.list.currentItem()
        if it is None:
            return None
        pid = it.data(Qt.UserRole)
        return str(pid) if pid else None

    def _on_open_original_clicked(self) -> None:
        pid = self._current_problem_id()
        if not pid:
            return
        try:
            p = self.hwp_restore.restore_to_temp_file(pid, prefix="preview_", suffix=".hwp")
            os.startfile(p)  # type: ignore[attr-defined]
        except Exception as e:
            QMessageBox.critical(self, "원본 HWP 보기", f"원본 HWP를 열 수 없습니다.\n\n{e}")

    def _on_finalize_clicked(self) -> None:
        d = self._draft
        if not d:
            return

        # Step2 입력값 반영 + 필수 검증(학습지명)
        self._apply_final_config_to_draft()
        if not (d.title or "").strip():
            QMessageBox.warning(self, "최종 생성", "학습지명을 입력해주세요.")
            try:
                self.title_input.setFocus()
            except Exception:
                pass
            return

        ids = self.current_problem_ids()
        if not ids:
            QMessageBox.warning(self, "최종 생성", "문항이 없습니다. (모두 삭제됨)")
            return

        default_name = (d.title or "worksheet").strip().replace("/", "_").replace("\\", "_")
        default_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{default_name}.hwp")
        out, _ = QFileDialog.getSaveFileName(self, "HWP 저장", default_path, "HWP Files (*.hwp)")
        if not out:
            return

        # 문항 순서대로 메타(단원/출처/난이도) 구성 → HWP 문항 아래 태그 삽입용
        problem_meta = [
            {
                "problem_id": pid,
                "unit": (self._unit.get(pid, "") or "").strip(),
                "source": (self._source.get(pid, "") or "").strip(),
                "difficulty": (self._diff.get(pid, "") or "").strip(),
            }
            for pid in ids
        ]
        tag_options = {
            "unit": bool(d.option_unit_tag),
            "source": bool(d.option_source_tag),
            "difficulty": bool(d.option_difficulty_tag),
        }

        try:
            composer = WorksheetHwpComposer(self.db)
            composer.compose(
                problem_ids=ids,
                output_path=out,
                title=d.title,
                teacher=d.creator,
                date_str=datetime.now().strftime("%Y.%m.%d"),
                tag_options=tag_options,
                problem_meta=problem_meta,
            )
        except WorksheetComposeError as e:
            QMessageBox.critical(self, "최종 생성 실패", str(e))
            return
        except Exception as e:
            QMessageBox.critical(self, "최종 생성 실패", f"예기치 못한 오류: {e}")
            return

        try:
            os.startfile(out)  # type: ignore[attr-defined]
        except Exception:
            pass

        numbered = [{"no": i + 1, "problem_id": pid} for i, pid in enumerate(ids)]
        # ✅ 후속 단계에서 쓰도록 문서 메타도 함께 전달(기존 구조 유지 + 확장)
        self.finalized.emit(
            {
                "output_path": out,
                "numbered": numbered,
                "problem_ids": list(ids),
                "title": d.title,
                "creator": d.creator,
                "grade": d.grade,
                "type_text": d.type_text,
                "option_unit_tag": bool(d.option_unit_tag),
                "option_source_tag": bool(d.option_source_tag),
                "option_difficulty_tag": bool(d.option_difficulty_tag),
            }
        )
        QMessageBox.information(self, "완료", "HWP가 생성되었습니다.")


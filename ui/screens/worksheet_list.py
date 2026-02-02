"""
학습지 목록 화면 (고밀도 Compact 대시보드 리스트 + 일괄 관리)

- Row 높이 65~70px (고정: 68px)
- 카드 간격 8px
- 단일 라인 레이아웃: [체크] [학년] [유형] [제목(최대)] [메타] [다운로드]
- 상단 Action Bar: 전체 선택 / 선택 삭제 / 일괄 다운로드
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Set

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtWidgets import (
    QAction,
    QButtonGroup,
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
    QGraphicsDropShadowEffect,
)

from database.sqlite_connection import SQLiteConnection
from database.repositories import WorksheetRepository, WorksheetAssignmentRepository
from processors.hwp.hwp_reader import HWPReader, HWPNotInstalledError, HWPInitializationError
from ui.components.standard_action_dialog import DialogAction, StandardActionDialog
from ui.components.student_select_dialog import StudentSelectDialog
from ui.components.standard_message import show_info, show_warning, confirm

try:
    import qtawesome as qta  # type: ignore

    _QTA_AVAILABLE = True
except Exception:
    qta = None
    _QTA_AVAILABLE = False


def _pick_font(size_pt: int, bold: bool = False, extra_bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    if extra_bold:
        f.setWeight(QFont.ExtraBold)
    elif bold:
        f.setBold(True)
    return f


def _qta_pixmap(icon_name: str, color: str, size: int) -> Optional[QPixmap]:
    if not _QTA_AVAILABLE or qta is None:
        return None
    try:
        return qta.icon(icon_name, color=color).pixmap(size, size)
    except Exception:
        return None


def _fallback_search_pixmap(color_hex: str = "#475569", size: int = 16) -> QPixmap:
    """qtawesome 미설치 환경에서도 쓰는 돋보기 픽스맵."""
    from PyQt5.QtGui import QPainter, QPen, QColor

    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    c = QColor(color_hex)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(c)
    pen.setWidthF(1.8)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    p.setPen(pen)
    p.setBrush(Qt.NoBrush)

    r = size * 0.36
    cx = size * 0.42
    cy = size * 0.42
    p.drawEllipse(int(cx - r), int(cy - r), int(2 * r), int(2 * r))
    # handle
    p.drawLine(int(cx + r * 0.55), int(cy + r * 0.55), int(size - 2), int(size - 2))
    p.end()
    return pm


@dataclass
class StudySheetItem:
    id: str
    grade: str
    type_text: str
    title: str
    date: str
    teacher: str
    has_hwp: bool = True
    has_pdf: bool = False


class CompactStudySheetRow(QFrame):
    selected_changed = pyqtSignal(str, bool)  # (item_id, selected)
    download_requested = pyqtSignal(str, str)  # (item_id, kind: "PDF"|"HWP")

    def __init__(self, item: StudySheetItem, selected: bool, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.item = item
        self.setObjectName("CompactRow")
        self.setFixedHeight(68)
        self.setFocusPolicy(Qt.NoFocus)
        self.setCursor(Qt.PointingHandCursor)

        self._is_checked = bool(selected)
        self._is_hovered = False

        # 카드 미세 그림자(blur 15, opacity ~0.06)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 15))
        self.setGraphicsEffect(shadow)

        layout = QHBoxLayout(self)
        # ✅ 카드 내부 좌/우 패딩을 과하지 않게 조정(제목이 길어도 우측 버튼과 겹치지 않도록 폭 확보)
        layout.setContentsMargins(18, 5, 18, 5)
        layout.setSpacing(14)

        # 1) 체크박스
        self.checkbox = QCheckBox()
        self.checkbox.setFocusPolicy(Qt.NoFocus)
        self.checkbox.setChecked(self._is_checked)
        self.checkbox.stateChanged.connect(self._on_checkbox_changed)
        layout.addWidget(self.checkbox, alignment=Qt.AlignVCenter)

        # 2) 학년 배지
        grade_text = (item.grade or "").strip()
        if not grade_text:
            grade_text = "—"
        self.grade_badge = QLabel(grade_text)
        self.grade_badge.setFixedHeight(22)
        self.grade_badge.setMinimumWidth(42)
        self.grade_badge.setAlignment(Qt.AlignCenter)
        self.grade_badge.setFont(_pick_font(9, bold=True))
        self.grade_badge.setFocusPolicy(Qt.NoFocus)
        try:
            if (item.grade or "").strip() == "":
                self.grade_badge.setToolTip("학년 미지정")
        except Exception:
            pass
        layout.addWidget(self.grade_badge, alignment=Qt.AlignVCenter)

        # 3) 유형 배지
        self.type_badge = QLabel(item.type_text)
        self.type_badge.setFixedHeight(22)
        self.type_badge.setAlignment(Qt.AlignCenter)
        self.type_badge.setFont(_pick_font(9, bold=True))
        self.type_badge.setFocusPolicy(Qt.NoFocus)
        self.type_badge.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        self.type_badge.setContentsMargins(10, 0, 10, 0)
        layout.addWidget(self.type_badge, alignment=Qt.AlignVCenter)

        # 4) 제목 (최대한 넓게)
        self.title_label = QLabel(item.title)
        self.title_label.setFont(_pick_font(11, bold=True))
        self.title_label.setFocusPolicy(Qt.NoFocus)
        self.title_label.setStyleSheet("color: #1E293B;")
        self.title_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.title_label.setToolTip(item.title)
        self.title_label.setMinimumWidth(240)
        layout.addWidget(self.title_label, 1)

        # 5) 메타 정보 (생성일/출제자 고정 너비 + 수직 중앙 정렬로 줄 맞춤)
        meta_wrap = QWidget()
        meta_wrap.setFocusPolicy(Qt.NoFocus)
        meta_layout = QHBoxLayout(meta_wrap)
        meta_layout.setContentsMargins(0, 0, 0, 0)
        meta_layout.setSpacing(15)
        meta_layout.setAlignment(Qt.AlignVCenter)

        def _create_info_item(icon_name: str, text: str, label_width: int = 0) -> QWidget:
            item_w = QWidget()
            item_w.setFocusPolicy(Qt.NoFocus)
            l = QHBoxLayout(item_w)
            l.setContentsMargins(0, 0, 0, 0)
            l.setSpacing(6)
            l.setAlignment(Qt.AlignVCenter)

            icon = QLabel()
            icon.setFocusPolicy(Qt.NoFocus)
            icon.setFixedSize(14, 14)
            pm = _qta_pixmap(icon_name, "#94A3B8", 14)
            if pm is not None:
                icon.setPixmap(pm)
            else:
                icon.setText("")

            label = QLabel(text)
            label.setFocusPolicy(Qt.NoFocus)
            f = _pick_font(9)
            f.setWeight(QFont.Medium)
            label.setFont(f)
            label.setObjectName("MetaText")
            label.setStyleSheet("color: #222222;")
            if label_width > 0:
                label.setFixedWidth(label_width)
                label.setAlignment(Qt.AlignCenter)

            l.addWidget(icon, alignment=Qt.AlignVCenter)
            l.addWidget(label, alignment=Qt.AlignVCenter)
            return item_w

        meta_layout.addWidget(_create_info_item("fa5s.calendar-alt", item.date, label_width=100))
        meta_layout.addWidget(_create_info_item("fa5s.user-tie", item.teacher, label_width=80))

        layout.addWidget(meta_wrap, alignment=Qt.AlignVCenter)

        # 6) 다운로드 버튼군 (동일 크기로 통일)
        self.btn_pdf = QPushButton("PDF")
        self.btn_pdf.setObjectName("ActionBtn")
        self.btn_hwp = QPushButton("HWP")
        self.btn_hwp.setObjectName("ActionBtn")
        for btn, kind in ((self.btn_pdf, "PDF"), (self.btn_hwp, "HWP")):
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFixedSize(50, 28)
            btn.setFont(_pick_font(8, bold=True))
            btn.clicked.connect(lambda checked=False, k=kind: self.download_requested.emit(self.item.id, k))

        # 파일 유무에 따라 버튼 활성/비활성
        try:
            self.btn_hwp.setEnabled(bool(getattr(item, "has_hwp", False)))
            # PDF는 (1) GridFS에 이미 있거나, (2) HWP가 있으면 온디맨드 변환으로 제공
            self.btn_pdf.setEnabled(bool(getattr(item, "has_pdf", False) or getattr(item, "has_hwp", False)))
            if not bool(getattr(item, "has_hwp", False)):
                self.btn_hwp.setToolTip("HWP 파일이 없습니다.")
            if not bool(getattr(item, "has_pdf", False)) and bool(getattr(item, "has_hwp", False)):
                self.btn_pdf.setToolTip("PDF는 HWP로부터 생성됩니다(변환 시간이 걸릴 수 있음).")
            elif not bool(getattr(item, "has_pdf", False)):
                self.btn_pdf.setToolTip("PDF 파일이 없습니다.")
        except Exception:
            pass
        layout.addWidget(self.btn_pdf, alignment=Qt.AlignVCenter)
        layout.addWidget(self.btn_hwp, alignment=Qt.AlignVCenter)

        # 행 전체 클릭 시 체크박스와 동일 효과: 클릭이 떨어지는 자식은 마우스 이벤트 투과
        for w in (self.grade_badge, self.type_badge, self.title_label, meta_wrap):
            w.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        for i in range(meta_layout.count()):
            it = meta_layout.itemAt(i)
            if it and it.widget():
                it.widget().setAttribute(Qt.WA_TransparentForMouseEvents, True)

        self._apply_styles()

    def mousePressEvent(self, event):  # noqa: N802 (Qt naming)
        if event.button() == Qt.LeftButton:
            self.set_selected(not self._is_checked, emit=True)
            event.accept()
            return
        super().mousePressEvent(event)

    def enterEvent(self, event):  # noqa: N802 (Qt naming)
        self._is_hovered = True
        self._apply_styles()
        super().enterEvent(event)

    def leaveEvent(self, event):  # noqa: N802 (Qt naming)
        self._is_hovered = False
        self._apply_styles()
        super().leaveEvent(event)

    def set_selected(self, selected: bool, *, emit: bool = False) -> None:
        selected = bool(selected)
        if self._is_checked == selected:
            return
        self._is_checked = selected
        self.checkbox.blockSignals(True)
        self.checkbox.setChecked(selected)
        self.checkbox.blockSignals(False)
        self._apply_styles()
        if emit:
            self.selected_changed.emit(self.item.id, selected)

    def is_selected(self) -> bool:
        return bool(self._is_checked)

    def _on_checkbox_changed(self, state: int) -> None:
        self._is_checked = state == Qt.Checked
        self._apply_styles()
        self.selected_changed.emit(self.item.id, self._is_checked)

    def _apply_styles(self) -> None:
        # 카드: 얇은 테두리만, 그림자 제거
        if self._is_checked:
            bg = "#EFF6FF"
        elif self._is_hovered:
            bg = "#F9FAFB"
        else:
            bg = "#FFFFFF"

        self.setStyleSheet(
            f"""
            QFrame#CompactRow {{
                background-color: {bg};
                border: 1px solid #E2E8F0;
                border-radius: 12px;
            }}
            /* 충돌 방지: Row 내부 스타일 초기화(override) */
            QFrame#CompactRow * {{
                border: none;
                outline: none;
                background: transparent;
            }}
            QPushButton:focus {{ outline: none; }}
            /* ✅ 행 선택 체크박스: 흰 배경에서도 보이도록 테두리/체크색 강제 */
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                background: #FFFFFF;
                border: 2px solid #94A3B8;
                border-radius: 4px;
            }}
            QCheckBox::indicator:hover {{
                border-color: #2563EB;
            }}
            QCheckBox::indicator:checked {{
                background: #2563EB;
                border-color: #2563EB;
            }}
            QLabel#MetaText {{
                color: #222222;
                font-size: 10pt;
                font-weight: 600;
                background: transparent;
                border: none;
            }}

            QPushButton#ActionBtn {{
                /* ✅ 전역 테마의 큰 padding 때문에 텍스트가 잘리는 문제 방지 */
                padding: 0px;
                margin: 0px;
                background: #F8FAFC;
                border: 1.5px solid #CBD5E1;
                color: #1E293B;
                font-weight: 700;
                font-size: 9pt;
                border-radius: 6px;
            }}
            QPushButton#ActionBtn:hover {{
                border-color: #2563EB;
                color: #2563EB;
                background: #F0F7FF;
            }}
            """
        )

        # 학년 배지 (compact)
        self.grade_badge.setStyleSheet(
            """
            QLabel {
                background-color: #EFF6FF;
                color: #2563EB;
                border-radius: 6px;
                border: 1px solid #DBEAFE;
                padding: 0px 8px;
            }
            """
        )

        # 유형 배지 (compact, 파스텔)
        self.type_badge.setStyleSheet(self._type_badge_style(self.item.type_text))

        # 버튼 스타일은 QSS(QPushButton#ActionBtn)에서 강제
        self.btn_pdf.setStyleSheet("")
        self.btn_hwp.setStyleSheet("")

    def _type_badge_style(self, t: str) -> str:
        if "내신" in t:
            bg, fg, bd = "#EEF2FF", "#4F46E5", "#E0E7FF"
        elif "교재" in t:
            bg, fg, bd = "#ECFDF5", "#059669", "#D1FAE5"
        else:
            bg, fg, bd = "#FDF4FF", "#A855F7", "#F5D0FE"
        return f"""
            QLabel {{
                background-color: {bg};
                color: {fg};
                border: 1px solid {bd};
                border-radius: 6px;
                padding: 0px 10px;
                font-weight: 700;
            }}
        """


class WorksheetListScreen(QWidget):
    """학습지 목록 화면"""

    create_requested = pyqtSignal()

    def __init__(self, db_connection: Optional[SQLiteConnection] = None, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.db_connection = db_connection
        self.ws_repo: Optional[WorksheetRepository] = WorksheetRepository(db_connection) if db_connection else None
        self.assign_repo: Optional[WorksheetAssignmentRepository] = (
            WorksheetAssignmentRepository(db_connection) if db_connection else None
        )
        self._items: List[StudySheetItem] = []
        self._selected_ids: Set[str] = set()
        self._visible_ids: List[str] = []
        self._current_grade_tab: str = "전체"
        self._select_all_updating = False
        self._page: int = 0
        self._page_size: int = 15

        self._init_ui()
        self.reload_from_db()

    def reload_from_db(self) -> None:
        """DB에서 학습지 목록을 새로 로드합니다."""
        self._items = []

        if not self.ws_repo or not self.db_connection:
            self.refresh_list()
            return

        if not self.db_connection.is_connected():
            # 오프라인이면 빈 목록으로 유지(목업을 보여주면 혼동될 수 있음)
            self._selected_ids.clear()
            self.refresh_list()
            return

        try:
            worksheets = self.ws_repo.list_all()
            items: List[StudySheetItem] = []
            for ws in worksheets:
                ws_id = str(ws.id or "")
                if not ws_id:
                    continue
                dt = ws.created_at
                date_str = dt.strftime("%Y.%m.%d") if dt else ""
                items.append(
                    StudySheetItem(
                        id=ws_id,
                        grade=(ws.grade or "").strip(),
                        type_text=(ws.type_text or "").strip(),
                        title=(ws.title or "").strip(),
                        date=date_str,
                        teacher=(ws.creator or "").strip(),
                        has_hwp=bool(getattr(ws, "hwp_file_id", None)),
                        has_pdf=bool(getattr(ws, "pdf_file_id", None)),
                    )
                )

            self._items = items
            # 선택 상태 정리(삭제된 id 제거)
            alive = {it.id for it in self._items}
            self._selected_ids = {sid for sid in self._selected_ids if sid in alive}
        except Exception:
            # 로드 실패 시에도 UI는 유지(빈 목록)
            self._items = []

        self.refresh_list()

    def _init_ui(self) -> None:
        self.setObjectName("worksheetList")
        self.setStyleSheet(
            """
            /* 충돌 방지: 화면 내부 스타일 초기화(override) */
            QWidget#worksheetList * {
                border: none;
                background: transparent;
                outline: none;
                font-family: 'Pretendard','Malgun Gothic','맑은 고딕';
            }
            QWidget#worksheetList {
                background-color: #F1F5F9;
            }
            QLabel, QFrame {
                border: none;
                outline: none;
            }
            QPushButton { outline: none; }
            QPushButton:focus { outline: none; }
            """
        )

        root = QVBoxLayout(self)
        # ✅ Sidebar(180px)와 본문 사이 '황금 여백' 확보 (약 25~35px)
        # - MainWindow는 sidebar와 stacked_widget 사이 spacing=0 이므로,
        #   이 화면의 left margin이 곧 "사이드바-본문 간격"이 됩니다.
        # ✅ 요구사항 업데이트:
        # - Left Pull: 사이드바-본문 간격을 15~20px로 축소
        # - Right Padding: 우측 안전 여백을 60~80px로 확대
        #   (버튼이 화면 베젤에 너무 붙어 보이지 않도록)
        root.setContentsMargins(18, 16, 72, 24)
        root.setSpacing(12)

        # 상단: 필터/검색/생성(기존 UX 유지)
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        self.grade_group = QButtonGroup(self)
        self.grade_group.setExclusive(True)
        self.grade_buttons: Dict[str, QPushButton] = {}
        for i, grade in enumerate(["전체", "초", "중", "고"]):
            btn = QPushButton(grade)
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setMinimumHeight(34)
            btn.setMinimumWidth(52)
            btn.setFont(_pick_font(10, bold=True))
            btn.setObjectName("FilterBtn")
            if grade == "전체":
                btn.setChecked(True)
            self.grade_group.addButton(btn, i)
            self.grade_buttons[grade] = btn
            top_bar.addWidget(btn)

        self.grade_group.buttonClicked.connect(self._on_grade_tab_clicked)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("검색...")
        self.search_input.setMinimumHeight(34)
        self.search_input.setMinimumWidth(320)
        self.search_input.setClearButtonEnabled(True)
        self.search_input.setFont(_pick_font(10))
        self.search_input.textChanged.connect(self.refresh_list)
        self.search_input.setStyleSheet(
            """
            QLineEdit {
                background-color: #FFFFFF;
                border: 1.5px solid #CBD5E1;
                border-radius: 8px;
                padding-left: 38px;  /* 아이콘 공간(겹침/잘림 방지) */
                padding-right: 12px;
                padding-top: 8px;
                padding-bottom: 8px;
                font-size: 11pt;
                color: #1E293B;  /* 입력 텍스트 */
                font-weight: 600;
            }
            QLineEdit::placeholder {
                color: #475569;
                font-weight: 600;
            }
            QLineEdit:focus {
                background-color: #FFFFFF;
                border: 2px solid #2563EB;
                padding-left: 37px;
                padding-right: 11px;
                padding-top: 7px;
                padding-bottom: 7px;
            }
            """
        )
        # 돋보기 아이콘 잘림 방지:
        # - 아이콘을 왼쪽 끝에서 약 12px 안쪽으로 넣기 위해 "스페이서 액션"을 먼저 추가합니다.
        try:
            from PyQt5.QtGui import QIcon

            # spacer(12px)
            spacer_pm = QPixmap(12, 16)
            spacer_pm.fill(Qt.transparent)
            spacer_act = QAction(QIcon(spacer_pm), "", self.search_input)
            spacer_act.setEnabled(False)
            self.search_input.addAction(spacer_act, QLineEdit.LeadingPosition)

            # icon(16px)
            if _QTA_AVAILABLE and qta is not None:
                icon_act = QAction(qta.icon("fa5s.search", color="#64748B"), "", self.search_input)
            else:
                icon_act = QAction(QIcon(_fallback_search_pixmap("#64748B", 16)), "", self.search_input)
            icon_act.setEnabled(False)
            self.search_input.addAction(icon_act, QLineEdit.LeadingPosition)
        except Exception:
            pass
        top_bar.addWidget(self.search_input)

        top_bar.addStretch(1)

        btn_create = QPushButton("학습지 생성")
        btn_create.setMinimumHeight(36)
        btn_create.setMinimumWidth(120)
        btn_create.setCursor(Qt.PointingHandCursor)
        btn_create.setFocusPolicy(Qt.NoFocus)
        btn_create.setFont(_pick_font(10, bold=True))
        btn_create.setObjectName("create")
        btn_create.clicked.connect(self.create_requested.emit)
        btn_create.setStyleSheet(
            """
            QPushButton#create {
                background-color: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 800;
                font-family: 'Pretendard','Segoe UI','Malgun Gothic','맑은 고딕';
            }
            QPushButton#create:hover {
                background-color: #1D4ED8;
            }
            """
        )
        top_bar.addWidget(btn_create)

        root.addLayout(top_bar)

        # 액션 바(고정)
        self.action_bar = QFrame()
        self.action_bar.setObjectName("actionBar")
        self.action_bar.setFixedHeight(44)
        self.action_bar.setStyleSheet(
            """
            QFrame#actionBar {
                background-color: #F8FAFC;
                border: 1px solid #F1F5F9;
                border-radius: 12px;
            }
            QLabel, QFrame, QPushButton, QCheckBox {
                border: none;
                outline: none;
                background: transparent;
            }
            /* ✅ 상단 '전체 선택' 체크박스도 대비 강화 */
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
                background: #FFFFFF;
                border: 2px solid #94A3B8;
                border-radius: 4px;
            }
            QCheckBox::indicator:hover {
                border-color: #2563EB;
            }
            QCheckBox::indicator:checked {
                background: #2563EB;
                border-color: #2563EB;
            }
            """
        )
        ab = QHBoxLayout(self.action_bar)
        ab.setContentsMargins(16, 0, 16, 0)
        ab.setSpacing(10)

        self.chk_select_all = QCheckBox("전체 선택")
        self.chk_select_all.setFont(_pick_font(10, bold=True))
        self.chk_select_all.setFocusPolicy(Qt.NoFocus)
        self.chk_select_all.setTristate(True)
        self.chk_select_all.stateChanged.connect(self._on_select_all_changed)
        ab.addWidget(self.chk_select_all, alignment=Qt.AlignVCenter)

        self.lbl_selected = QLabel("선택 0")
        self.lbl_selected.setFont(_pick_font(9))
        self.lbl_selected.setStyleSheet("color: #475569; font-weight: 600;")
        ab.addWidget(self.lbl_selected, alignment=Qt.AlignVCenter)

        ab.addStretch(1)

        self.btn_bulk_download = QPushButton("일괄 다운로드")
        self.btn_bulk_download.setFocusPolicy(Qt.NoFocus)
        self.btn_bulk_download.setCursor(Qt.PointingHandCursor)
        self.btn_bulk_download.setFont(_pick_font(9, bold=True))
        self.btn_bulk_download.setFixedHeight(30)
        self.btn_bulk_download.clicked.connect(self._on_bulk_download_clicked)
        self.btn_bulk_download.setStyleSheet(
            """
            QPushButton {
                background-color: #DBEAFE;
                color: #1E40AF;
                border: none;
                border-radius: 10px;
                padding: 6px 12px;
                font-weight: 800;
                font-family: 'Pretendard','Segoe UI','Malgun Gothic','맑은 고딕';
            }
            QPushButton:hover {
                background-color: #BFDBFE;
                color: #1D4ED8;
            }
            QPushButton:disabled {
                background-color: #E5E7EB;
                color: #94A3B8;
            }
            """
        )
        ab.addWidget(self.btn_bulk_download, alignment=Qt.AlignVCenter)

        self.btn_assign = QPushButton("출제하기")
        self.btn_assign.setFocusPolicy(Qt.NoFocus)
        self.btn_assign.setCursor(Qt.PointingHandCursor)
        self.btn_assign.setFont(_pick_font(9, bold=True))
        self.btn_assign.setFixedHeight(30)
        self.btn_assign.clicked.connect(self._on_assign_clicked)
        self.btn_assign.setStyleSheet(
            """
            QPushButton {
                background-color: #DCFCE7;
                color: #166534;
                border: none;
                border-radius: 10px;
                padding: 6px 12px;
                font-weight: 800;
                font-family: 'Pretendard','Segoe UI','Malgun Gothic','맑은 고딕';
            }
            QPushButton:hover {
                background-color: #BBF7D0;
                color: #14532D;
            }
            QPushButton:disabled {
                background-color: #E5E7EB;
                color: #94A3B8;
            }
            """
        )
        ab.addWidget(self.btn_assign, alignment=Qt.AlignVCenter)

        self.btn_bulk_delete = QPushButton("선택 삭제")
        self.btn_bulk_delete.setFocusPolicy(Qt.NoFocus)
        self.btn_bulk_delete.setCursor(Qt.PointingHandCursor)
        self.btn_bulk_delete.setFont(_pick_font(9, bold=True))
        self.btn_bulk_delete.setFixedHeight(30)
        self.btn_bulk_delete.clicked.connect(self._on_bulk_delete_clicked)
        self.btn_bulk_delete.setStyleSheet(
            """
            QPushButton {
                background-color: #FEE2E2;
                color: #991B1B;
                border: none;
                border-radius: 10px;
                padding: 6px 12px;
                font-weight: 800;
                font-family: 'Pretendard','Segoe UI','Malgun Gothic','맑은 고딕';
            }
            QPushButton:hover {
                background-color: #FECACA;
                color: #7F1D1D;
            }
            QPushButton:disabled {
                background-color: #E5E7EB;
                color: #94A3B8;
            }
            """
        )
        ab.addWidget(self.btn_bulk_delete, alignment=Qt.AlignVCenter)

        root.addWidget(self.action_bar)

        # 리스트(스크롤)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        self.list_container = QWidget()
        self.list_layout = QVBoxLayout(self.list_container)
        self.list_layout.setContentsMargins(0, 0, 0, 0)
        self.list_layout.setSpacing(8)  # ✅ 고밀도 간격
        self.list_layout.addStretch(1)
        self.scroll.setWidget(self.list_container)

        root.addWidget(self.scroll)

        # 페이지네이션 바 (한 페이지에 15개)
        pagination_bar = QFrame()
        pagination_bar.setObjectName("PaginationBar")
        pagination_bar.setFixedHeight(40)
        pagination_bar.setStyleSheet(
            """
            QFrame#PaginationBar {
                background-color: transparent;
                border: none;
            }
            QPushButton#PaginationBtn {
                background-color: #F1F5F9;
                color: #334155;
                border: 1px solid #E2E8F0;
                border-radius: 8px;
                padding: 6px 14px;
                font-weight: 700;
            }
            QPushButton#PaginationBtn:hover:enabled {
                background-color: #E2E8F0;
                color: #1E293B;
            }
            QPushButton#PaginationBtn:disabled {
                background-color: #F8FAFC;
                color: #94A3B8;
            }
            """
        )
        pagination_layout = QHBoxLayout(pagination_bar)
        pagination_layout.setContentsMargins(0, 4, 0, 0)
        pagination_layout.setSpacing(10)

        self.lbl_pagination = QLabel("0-0 / 0건")
        self.lbl_pagination.setFont(_pick_font(9, bold=True))
        self.lbl_pagination.setStyleSheet("color: #475569;")
        pagination_layout.addWidget(self.lbl_pagination, alignment=Qt.AlignVCenter)

        pagination_layout.addStretch(1)

        self.btn_prev_page = QPushButton("이전")
        self.btn_prev_page.setObjectName("PaginationBtn")
        self.btn_prev_page.setFocusPolicy(Qt.NoFocus)
        self.btn_prev_page.setCursor(Qt.PointingHandCursor)
        self.btn_prev_page.setFixedHeight(32)
        self.btn_prev_page.clicked.connect(self._on_prev_page)
        pagination_layout.addWidget(self.btn_prev_page, alignment=Qt.AlignVCenter)

        self.btn_next_page = QPushButton("다음")
        self.btn_next_page.setObjectName("PaginationBtn")
        self.btn_next_page.setFocusPolicy(Qt.NoFocus)
        self.btn_next_page.setCursor(Qt.PointingHandCursor)
        self.btn_next_page.setFixedHeight(32)
        self.btn_next_page.clicked.connect(self._on_next_page)
        pagination_layout.addWidget(self.btn_next_page, alignment=Qt.AlignVCenter)

        root.addWidget(pagination_bar)

        self._sync_action_bar_state()

        # 상단 탭 버튼 스타일(간단)
        # 상단 필터 버튼 가독성(덮어쓰기)
        for btn in self.grade_buttons.values():
            btn.setStyleSheet(
                """
                QPushButton#FilterBtn {
                    background-color: #F1F5F9;
                    color: #1E293B;
                    font-weight: 600;
                    border-radius: 8px;
                    padding: 8px 16px;
                }
                QPushButton#FilterBtn:hover {
                    background-color: #E2E8F0;
                }
                QPushButton#FilterBtn:checked {
                    background-color: #2563EB;
                    color: #FFFFFF;
                }
                """
            )

    def _load_mock_data(self) -> None:
        self._items = [
            StudySheetItem("ws-001", "고2", "내신기출", "지수로그함수 내신기출 복습", "2025.12.22", "이창현T"),
            StudySheetItem("ws-002", "고3", "시중교재", "251225(월) 쎈수학 문항복습", "2025.12.25", "김가나T"),
            StudySheetItem("ws-003", "중1", "통합", "지수로그함수 내신기출 복습", "2025.12.28", "나다라T"),
            StudySheetItem("ws-004", "고2", "내신기출", "251231(수) 쎈수학 문항복습", "2025.12.31", "강강강T"),
            StudySheetItem("ws-005", "고1", "시중교재", "지수로그함수 내신기출 복습", "2026.01.02", "이창현T"),
        ]

    def _on_grade_tab_clicked(self, btn: QPushButton) -> None:
        self._current_grade_tab = btn.text().strip()
        self.refresh_list()

    def _grade_match(self, grade: str) -> bool:
        tab = self._current_grade_tab
        if tab == "전체":
            return True
        if tab == "초":
            return grade.startswith("초")
        if tab == "중":
            return grade.startswith("중")
        if tab == "고":
            return grade.startswith("고")
        return True

    def refresh_list(self) -> None:
        # 기존 row 제거(stretch 제외)
        while self.list_layout.count() > 1:
            it = self.list_layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.setParent(None)

        query = (self.search_input.text() or "").strip().lower()

        visible_items: List[StudySheetItem] = []
        for it in self._items:
            if not self._grade_match(it.grade):
                continue
            if query:
                hay = f"{it.grade} {it.type_text} {it.title} {it.date} {it.teacher}".lower()
                if query not in hay:
                    continue
            visible_items.append(it)

        self._visible_ids = [it.id for it in visible_items]
        total = len(visible_items)
        total_pages = max(1, (total + self._page_size - 1) // self._page_size)
        self._page = min(self._page, max(0, total_pages - 1))
        start = self._page * self._page_size
        page_items = visible_items[start : start + self._page_size]

        if not visible_items:
            empty = QLabel("표시할 학습지가 없습니다.")
            empty.setFont(_pick_font(10, bold=True))
            empty.setStyleSheet("color: #475569;")
            empty.setAlignment(Qt.AlignCenter)
            empty.setMinimumHeight(160)
            self.list_layout.insertWidget(0, empty)
            self.lbl_pagination.setText("0-0 / 0건")
            self.btn_prev_page.setEnabled(False)
            self.btn_next_page.setEnabled(False)
            self._sync_action_bar_state()
            return

        for it in page_items:
            row = CompactStudySheetRow(it, selected=it.id in self._selected_ids)
            row.selected_changed.connect(self._on_row_selected_changed)
            row.download_requested.connect(self._on_row_download_requested)
            self.list_layout.insertWidget(self.list_layout.count() - 1, row)

        end = start + len(page_items)
        self.lbl_pagination.setText(f"{start + 1}-{end} / {total}건")
        self.btn_prev_page.setEnabled(self._page > 0)
        self.btn_next_page.setEnabled(self._page < total_pages - 1)
        self._sync_action_bar_state()

    def _on_row_selected_changed(self, item_id: str, selected: bool) -> None:
        if selected:
            self._selected_ids.add(item_id)
        else:
            self._selected_ids.discard(item_id)
        self._sync_action_bar_state()

    def _on_prev_page(self) -> None:
        if self._page > 0:
            self._page -= 1
            self.refresh_list()

    def _on_next_page(self) -> None:
        total = len(self._visible_ids)
        total_pages = max(1, (total + self._page_size - 1) // self._page_size)
        if self._page < total_pages - 1:
            self._page += 1
            self.refresh_list()

    def _on_assign_clicked(self) -> None:
        try:
            if not self.assign_repo or not self.db_connection:
                show_warning(self, "출제하기", "DB 연결이 설정되어 있지 않습니다.")
                return
            if not self.db_connection.is_connected():
                show_warning(self, "출제하기", "DB에 연결할 수 없습니다.")
                return

            ws_ids = sorted(list(self._selected_ids))
            if not ws_ids:
                show_info(self, "출제하기", "먼저 출제할 학습지를 선택하세요.")
                return

            dlg = StudentSelectDialog(db_connection=self.db_connection, parent=self)
            if dlg.exec_() != dlg.Accepted:
                return

            student_ids = dlg.selected_student_ids()
            if not student_ids:
                show_info(self, "출제하기", "출제할 학생이 선택되지 않았습니다.")
                return

            stats = self.assign_repo.assign_many(worksheet_ids=ws_ids, student_ids=student_ids, assigned_by="")
            inserted = int(stats.get("inserted", 0) or 0)
            skipped = int(stats.get("skipped", 0) or 0)

            show_info(
                self,
                "출제 완료",
                f"학습지 {len(ws_ids)}개를 학생 {len(student_ids)}명에게 출제했습니다.\n\n"
                f"- 신규 출제: {inserted}\n"
                f"- 이미 출제됨(스킵): {skipped}",
            )
        except Exception as e:
            show_warning(
                self,
                "출제하기 오류",
                f"출제 처리 중 오류가 발생했습니다.\n\n{type(e).__name__}: {e}",
            )

    def _safe_filename(self, name: str, *, fallback: str = "worksheet") -> str:
        s = (name or "").strip() or fallback
        # Windows에서 파일명으로 위험한 문자 제거
        bad = '<>:"/\\|?*'
        for ch in bad:
            s = s.replace(ch, "_")
        s = " ".join(s.split())
        return s[:120]  # 너무 긴 파일명 방지

    def _ascii_temp_dir(self) -> str:
        base = os.environ.get("SystemDrive", "C:") + os.sep
        d = os.path.join(base, "CH_LMS_TMP", "worksheet_exports")
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        return d

    def _on_row_download_requested(self, item_id: str, kind: str) -> None:
        if not self.ws_repo or not self.db_connection:
            show_warning(self, "다운로드", "DB 연결이 설정되어 있지 않습니다.")
            return
        if not self.db_connection.is_connected():
            show_warning(self, "다운로드", "DB에 연결할 수 없습니다.")
            return

        ws = self.ws_repo.find_by_id(item_id)
        if not ws:
            show_warning(self, "다운로드", "학습지 정보를 찾을 수 없습니다.")
            return

        k = (kind or "").strip().upper()
        title = (ws.title or "").strip() or "worksheet"

        # 1) HWP 다운로드: GridFS 바이트 → 저장
        if k == "HWP":
            hwp_bytes = self.ws_repo.get_file_bytes(ws, "HWP")
            if not hwp_bytes:
                show_warning(self, "다운로드", "HWP 파일이 없습니다.")
                return

            # ✅ 요구사항: HWP 버튼 클릭 시 "저장/열기" 선택 팝업 제공
            dlg = StandardActionDialog(
                parent=self,
                title="HWP",
                message="HWP 작업을 선택하세요.",
                actions=[
                    DialogAction(key="save", label="저장", is_primary=False),
                    DialogAction(key="open", label="열기", is_primary=True),
                    DialogAction(key="cancel", label="취소", is_primary=False),
                ],
                min_width=360,
            )
            dlg.exec_()

            if dlg.selected_key == "save":
                default_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{self._safe_filename(title)}.hwp")
                out, _ = QFileDialog.getSaveFileName(self, "HWP 저장", default_path, "HWP Files (*.hwp)")
                if not out:
                    return
                try:
                    with open(out, "wb") as f:
                        f.write(hwp_bytes)
                    show_info(self, "완료", "HWP가 저장되었습니다.")
                except Exception as e:
                    show_warning(self, "다운로드 실패", f"HWP 저장에 실패했습니다.\n\n{e}")
            elif dlg.selected_key == "open":
                self._open_hwp_bytes(ws, hwp_bytes)
            return

        # 2) PDF 다운로드
        if k == "PDF":
            dlg = StandardActionDialog(
                parent=self,
                title="PDF",
                message="PDF 작업을 선택하세요.",
                actions=[
                    DialogAction(key="save", label="저장", is_primary=False),
                    DialogAction(key="open", label="열기", is_primary=True),
                    DialogAction(key="cancel", label="취소", is_primary=False),
                ],
                min_width=360,
            )
            dlg.exec_()

            # 2-A) PDF 열기(저장 없이 바로 열기)
            if dlg.selected_key == "open":
                # 2-A-1) PDF가 이미 GridFS에 있으면 임시 파일로 열기
                pdf_bytes = self.ws_repo.get_file_bytes(ws, "PDF")
                if pdf_bytes:
                    self._open_pdf_bytes(ws, pdf_bytes)
                    return

                # 2-A-2) 없으면 HWP로부터 임시 PDF 생성 후 열기
                hwp_bytes = self.ws_repo.get_file_bytes(ws, "HWP")
                if not hwp_bytes:
                    show_warning(self, "PDF 열기", "HWP 파일이 없어 PDF를 만들 수 없습니다.")
                    return

                tmp_dir = self._ascii_temp_dir()
                fd, tmp_hwp = tempfile.mkstemp(prefix="ws_", suffix=".hwp", dir=tmp_dir)
                os.close(fd)
                try:
                    with open(tmp_hwp, "wb") as f:
                        f.write(hwp_bytes)
                except Exception as e:
                    try:
                        os.remove(tmp_hwp)
                    except Exception:
                        pass
                    show_warning(self, "PDF 열기", f"임시 HWP 파일 생성에 실패했습니다.\n\n{e}")
                    return

                # 임시 PDF 출력 경로(ASCII 짧은 경로)
                title2 = (ws.title or "").strip() or "worksheet"
                safe = self._safe_filename(f"{ws.grade}_{ws.type_text}_{title2}_{ws.creator}".strip("_ "), fallback="worksheet")
                temp_pdf = os.path.join(self._ascii_temp_open_dir(), f"{safe}.pdf")

                try:
                    with HWPReader() as reader:
                        opened = reader.open_document(tmp_hwp)
                        if not opened:
                            raise RuntimeError("한글(HWP)에서 파일을 열 수 없습니다.")
                        ok = reader.export_pdf(temp_pdf)
                        try:
                            reader.close_document()
                        except Exception:
                            pass
                        if not ok:
                            raise RuntimeError("PDF로 내보내기에 실패했습니다. (한글 버전/환경 이슈 가능)")
                    try:
                        os.startfile(temp_pdf)  # type: ignore[attr-defined]
                    except Exception as e:
                        show_warning(self, "열기 실패", f"PDF를 열 수 없습니다.\n\n{e}")
                except (HWPNotInstalledError, HWPInitializationError) as e:
                    show_warning(self, "PDF 열기", str(e))
                except Exception as e:
                    show_warning(self, "PDF 열기", f"PDF 생성/열기에 실패했습니다.\n\n{e}")
                finally:
                    try:
                        if os.path.exists(tmp_hwp):
                            os.remove(tmp_hwp)
                    except Exception:
                        pass
                return

            # 2-B) PDF 저장(파일로 저장)
            if dlg.selected_key == "save":
                default_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{self._safe_filename(title)}.pdf")
                out, _ = QFileDialog.getSaveFileName(self, "PDF 저장", default_path, "PDF Files (*.pdf)")
                if not out:
                    return

                # 2-B-1) PDF가 이미 GridFS에 있으면 그대로 저장
                pdf_bytes = self.ws_repo.get_file_bytes(ws, "PDF")
                if pdf_bytes:
                    try:
                        with open(out, "wb") as f:
                            f.write(pdf_bytes)
                        show_info(self, "완료", "PDF가 저장되었습니다.")
                    except Exception as e:
                        show_warning(self, "다운로드 실패", f"PDF 저장에 실패했습니다.\n\n{e}")
                    return

                # 2-B-2) 없으면 HWP로부터 온디맨드 변환
                hwp_bytes = self.ws_repo.get_file_bytes(ws, "HWP")
                if not hwp_bytes:
                    show_warning(self, "PDF 생성", "HWP 파일이 없어 PDF를 만들 수 없습니다.")
                    return

                tmp_dir = self._ascii_temp_dir()
                fd, tmp_hwp = tempfile.mkstemp(prefix="ws_", suffix=".hwp", dir=tmp_dir)
                os.close(fd)
                try:
                    with open(tmp_hwp, "wb") as f:
                        f.write(hwp_bytes)
                except Exception as e:
                    try:
                        os.remove(tmp_hwp)
                    except Exception:
                        pass
                    show_warning(self, "PDF 생성", f"임시 HWP 파일 생성에 실패했습니다.\n\n{e}")
                    return

                try:
                    with HWPReader() as reader:
                        opened = reader.open_document(tmp_hwp)
                        if not opened:
                            raise RuntimeError("한글(HWP)에서 파일을 열 수 없습니다.")
                        ok = reader.export_pdf(out)
                        try:
                            reader.close_document()
                        except Exception:
                            pass
                        if not ok:
                            raise RuntimeError("PDF로 내보내기에 실패했습니다. (한글 버전/환경 이슈 가능)")
                    show_info(self, "완료", "PDF가 저장되었습니다.")
                except (HWPNotInstalledError, HWPInitializationError) as e:
                    show_warning(self, "PDF 생성", str(e))
                except Exception as e:
                    show_warning(self, "PDF 생성", f"PDF 생성에 실패했습니다.\n\n{e}")
                finally:
                    try:
                        if os.path.exists(tmp_hwp):
                            os.remove(tmp_hwp)
                    except Exception:
                        pass
                return

            # cancel / close
            return

        show_warning(self, "다운로드", f"지원하지 않는 형식입니다: {kind}")

    def _ascii_temp_open_dir(self) -> str:
        base = os.environ.get("SystemDrive", "C:") + os.sep
        d = os.path.join(base, "CH_LMS_TMP", "worksheet_open")
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
        return d

    def _open_hwp_bytes(self, ws, hwp_bytes: bytes) -> None:
        """HWP 저장 없이 바로 열기(임시 파일 복원 후 os.startfile)."""
        title = (ws.title or "").strip() or "worksheet"
        safe = self._safe_filename(f"{ws.grade}_{ws.type_text}_{title}_{ws.creator}".strip("_ "), fallback="worksheet")
        temp_dir = self._ascii_temp_open_dir()
        temp_path = os.path.join(temp_dir, f"{safe}.hwp")

        try:
            with open(temp_path, "wb") as f:
                f.write(hwp_bytes)
        except Exception as e:
            show_warning(self, "열기 실패", f"임시 HWP 파일을 만들 수 없습니다.\n\n{e}")
            return

        try:
            os.startfile(temp_path)  # type: ignore[attr-defined]
        except Exception as e:
            show_warning(self, "열기 실패", f"HWP를 열 수 없습니다.\n\n{e}")

    def _open_pdf_bytes(self, ws, pdf_bytes: bytes) -> None:
        """PDF 저장 없이 바로 열기(임시 파일 복원 후 os.startfile)."""
        title = (ws.title or "").strip() or "worksheet"
        safe = self._safe_filename(f"{ws.grade}_{ws.type_text}_{title}_{ws.creator}".strip("_ "), fallback="worksheet")
        temp_dir = self._ascii_temp_open_dir()
        temp_path = os.path.join(temp_dir, f"{safe}.pdf")

        try:
            with open(temp_path, "wb") as f:
                f.write(pdf_bytes)
        except Exception as e:
            show_warning(self, "열기 실패", f"임시 PDF 파일을 만들 수 없습니다.\n\n{e}")
            return

        try:
            os.startfile(temp_path)  # type: ignore[attr-defined]
        except Exception as e:
            show_warning(self, "열기 실패", f"PDF를 열 수 없습니다.\n\n{e}")

    def _set_visible_rows_selected(self, selected: bool) -> None:
        # 현재 렌더된 row 위젯들만 찾아 체크 반영
        for i in range(self.list_layout.count()):
            w = self.list_layout.itemAt(i).widget()
            if isinstance(w, CompactStudySheetRow):
                w.set_selected(selected, emit=False)

        if selected:
            self._selected_ids.update(self._visible_ids)
        else:
            for vid in list(self._visible_ids):
                self._selected_ids.discard(vid)

        self._sync_action_bar_state()

    def _on_select_all_changed(self, state: int) -> None:
        if self._select_all_updating:
            return

        # 사용자 조작 기준: Checked -> 전체 선택, Unchecked -> 전체 해제
        if state == Qt.Checked:
            self._set_visible_rows_selected(True)
        elif state == Qt.Unchecked:
            self._set_visible_rows_selected(False)
        else:
            # PartiallyChecked는 내부 동기화용 상태로만 사용
            pass

    def _sync_action_bar_state(self) -> None:
        # 선택 수(전체 기준)
        selected_count = len(self._selected_ids)
        self.lbl_selected.setText(f"선택 {selected_count}")
        self.btn_bulk_delete.setEnabled(selected_count > 0)
        self.btn_bulk_download.setEnabled(selected_count > 0)
        try:
            self.btn_assign.setEnabled(selected_count > 0)
        except Exception:
            pass

        # select-all(현재 보이는 목록 기준)
        visible = [vid for vid in self._visible_ids]
        if not visible:
            state = Qt.Unchecked
        else:
            sel_visible = sum(1 for vid in visible if vid in self._selected_ids)
            if sel_visible == 0:
                state = Qt.Unchecked
            elif sel_visible == len(visible):
                state = Qt.Checked
            else:
                state = Qt.PartiallyChecked

        self._select_all_updating = True
        self.chk_select_all.setCheckState(state)
        self._select_all_updating = False

    def _on_bulk_delete_clicked(self) -> None:
        if not self._selected_ids:
            return

        ok = confirm(self, "선택 삭제", f"선택된 학습지 {len(self._selected_ids)}개를 삭제할까요?", ok_label="삭제", cancel_label="취소")
        if not ok:
            return

        if not self.ws_repo or not self.db_connection or not self.db_connection.is_connected():
            show_warning(self, "선택 삭제", "DB에 연결할 수 없습니다.")
            return

        ids = list(self._selected_ids)
        deleted = 0
        failed = 0
        for wid in ids:
            try:
                if self.ws_repo.delete(wid):
                    deleted += 1
                else:
                    failed += 1
            except Exception:
                failed += 1

        self._selected_ids.clear()
        self.reload_from_db()
        if failed:
            show_warning(self, "부분 완료", f"삭제 완료: {deleted}개\n삭제 실패: {failed}개")
        else:
            show_info(self, "완료", f"{deleted}개 학습지를 삭제했습니다.")

    def _on_bulk_download_clicked(self) -> None:
        if not self._selected_ids:
            return
        if not self.ws_repo or not self.db_connection or not self.db_connection.is_connected():
            show_warning(self, "일괄 다운로드", "DB에 연결할 수 없습니다.")
            return

        out_dir = QFileDialog.getExistingDirectory(self, "다운로드 폴더 선택", os.path.expanduser("~"))
        if not out_dir:
            return

        saved = 0
        skipped = 0

        # HWP만 일괄 다운로드( PDF는 개별 버튼에서 필요 시 변환)
        for wid in list(self._selected_ids):
            ws = self.ws_repo.find_by_id(wid)
            if not ws:
                skipped += 1
                continue
            hwp_bytes = self.ws_repo.get_file_bytes(ws, "HWP")
            if not hwp_bytes:
                skipped += 1
                continue

            title = (ws.title or "").strip() or "worksheet"
            fname = self._safe_filename(f"{ws.grade}_{ws.type_text}_{title}_{ws.creator}".strip("_ "), fallback="worksheet")
            out_path = os.path.join(out_dir, f"{fname}.hwp")
            try:
                with open(out_path, "wb") as f:
                    f.write(hwp_bytes)
                saved += 1
            except Exception:
                skipped += 1

        show_info(self, "일괄 다운로드", f"저장: {saved}개\n건너뜀/실패: {skipped}개")


# 외부/문서 호환 alias
StudySheetList = WorksheetListScreen


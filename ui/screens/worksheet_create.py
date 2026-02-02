"""
학습지 생성 화면

학습지 생성 폼 화면
"""
from __future__ import annotations

import traceback
from typing import List, Optional, Callable, Tuple

from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QLineEdit,
    QCheckBox,
    QRadioButton,
    QButtonGroup,
    QSlider,
    QTreeWidget,
    QTreeWidgetItem,
    QScrollArea,
    QMessageBox,
    QFrame,
    QSpinBox,
    QSizePolicy,
    QFormLayout,
    QGraphicsDropShadowEffect,
    QLayout,
    QLayoutItem,
    QWidgetItem,
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize, QRect, QPoint, QEvent, QTimer
from PyQt5.QtGui import QFont, QIntValidator, QColor, QPalette, QShowEvent

from core.unit_catalog import UNIT_CATALOG
from database.sqlite_connection import SQLiteConnection
from database.repositories import TextbookRepository, ExamRepository
from services.worksheet import UnitKey
from services.worksheet.worksheet_service import SelectedSources
from services.worksheet import WorksheetService, OrderOptions
from ui.components.source_select_dialogs import TextbookMultiSelectDialog, ExamMultiSelectDialog


class FlowLayout(QLayout):
    """
    간단 FlowLayout (칩/태그용)
    - 가로로 배치하다가 폭이 넘치면 다음 줄로 래핑
    """

    def __init__(self, parent: Optional[QWidget] = None, *, spacing: int = 8):
        super().__init__(parent)
        self._items: List[QLayoutItem] = []
        self._spacing = int(spacing)

    def addItem(self, item: QLayoutItem) -> None:  # noqa: N802 (Qt naming)
        self._items.append(item)

    def addWidget(self, w: QWidget) -> None:  # type: ignore[override]
        # QLayout.addWidget()는 내부적으로 addChildWidget()로 부모를 세팅합니다.
        # 커스텀 레이아웃에서도 동일하게 처리하지 않으면 위젯이 화면에 보이지 않을 수 있습니다.
        try:
            # PyQt에서 제공되는 경우(대부분) 사용
            self.addChildWidget(w)  # type: ignore[attr-defined]
        except Exception:
            try:
                pw = self.parentWidget()
                if pw is not None and w.parent() is not pw:
                    w.setParent(pw)
            except Exception:
                pass
        self.addItem(QWidgetItem(w))

    def count(self) -> int:  # noqa: N802
        return len(self._items)

    def itemAt(self, index: int) -> Optional[QLayoutItem]:  # noqa: N802
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def takeAt(self, index: int) -> Optional[QLayoutItem]:  # noqa: N802
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self) -> Qt.Orientations:  # noqa: N802
        return Qt.Orientations(0)

    def hasHeightForWidth(self) -> bool:  # noqa: N802
        return True

    def heightForWidth(self, width: int) -> int:  # noqa: N802
        return self._do_layout(QRect(0, 0, int(width), 0), test_only=True)

    def setGeometry(self, rect: QRect) -> None:  # noqa: N802
        super().setGeometry(rect)
        self._do_layout(rect, test_only=False)

    def sizeHint(self) -> QSize:  # noqa: N802
        return self.minimumSize()

    def minimumSize(self) -> QSize:  # noqa: N802
        s = QSize(0, 0)
        for it in self._items:
            s = s.expandedTo(it.minimumSize())
        m = self.contentsMargins()
        s += QSize(m.left() + m.right(), m.top() + m.bottom())
        return s

    def _do_layout(self, rect: QRect, *, test_only: bool) -> int:
        x = rect.x()
        y = rect.y()
        line_h = 0

        m = self.contentsMargins()
        x0 = x + m.left()
        y0 = y + m.top()
        x = x0
        y = y0
        effective_w = rect.width() - (m.left() + m.right())

        for it in self._items:
            hint = it.sizeHint()
            w = hint.width()
            h = hint.height()
            if x > x0 and (x - x0 + w) > effective_w:
                x = x0
                y += line_h + self._spacing
                line_h = 0
            if not test_only:
                it.setGeometry(QRect(QPoint(x, y), hint))
            x += w + self._spacing
            line_h = max(line_h, h)

        return (y - y0) + line_h + m.top() + m.bottom()


def _pick_font(size_pt: int, *, bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    if bold:
        f.setBold(True)
    else:
        try:
            f.setWeight(QFont.Medium)
        except Exception:
            pass
    return f


class WorksheetCreateScreen(QWidget):
    """학습지 생성 화면"""
    
    # 시그널 정의
    close_requested = pyqtSignal()  # 닫기 요청
    preview_requested = pyqtSignal(dict)  # 문항 편집 화면으로 이동 요청(payload)
    
    def __init__(self, db_connection: SQLiteConnection, parent=None):
        super().__init__(parent)
        self.db_connection = db_connection
        self.textbook_repo = TextbookRepository(db_connection)
        self.exam_repo = ExamRepository(db_connection)

        self.selected_textbook_ids: List[str] = []
        self.selected_exam_ids: List[str] = []

        self._unit_tree: Optional[QTreeWidget] = None
        # 출처 UI(칩 시스템)
        self._source_mode: str = "textbook"  # "textbook" | "exam"
        self._source_search: Optional[QLineEdit] = None
        self._result_flow: Optional[FlowLayout] = None
        self._selected_tb_flow: Optional[FlowLayout] = None
        self._selected_ex_flow: Optional[FlowLayout] = None

        self.worksheet_service = WorksheetService(db_connection)
        self._last_selected_problem_ids: List[str] = []
        self._saved_state_for_restore: Optional[dict] = None  # 미리보기에서 돌아올 때 복원용

        self.init_ui()

    def showEvent(self, event: QShowEvent) -> None:
        """화면이 표시될 때: 미리보기에서 돌아온 경우 저장된 상태 복원, 그 외에는 리셋."""
        super().showEvent(event)
        try:
            if self._saved_state_for_restore:
                self._restore_state(self._saved_state_for_restore)
            else:
                self.selected_textbook_ids = []
                self.selected_exam_ids = []
                self._clear_unit_selection()
                self.refresh_selected_sources_view()
                self._refresh_source_chips()
        except Exception:
            pass

    def _on_close_clicked(self) -> None:
        """닫기 클릭: 복원용 저장 상태 초기화 후 닫기(다음에 목록에서 생성 진입 시 빈 상태로)."""
        self._saved_state_for_restore = None
        self.close_requested.emit()

    def _clear_unit_selection(self) -> None:
        """단원 트리 전체 체크 해제."""
        tree = self._unit_tree
        if tree is None:
            return
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            s_item = root.child(i)
            for j in range(s_item.childCount()):
                m_item = s_item.child(j)
                for k in range(m_item.childCount()):
                    sub_item = m_item.child(k)
                    try:
                        sub_item.setCheckState(0, Qt.Unchecked)
                    except Exception:
                        pass

    def _set_unit_selection(self, unit_keys: List[UnitKey]) -> None:
        """단원 트리에서 주어진 UnitKey 목록에 해당하는 소단원만 체크하고 부모 노드 펼침."""
        tree = self._unit_tree
        if tree is None or not unit_keys:
            return
        unit_set = {(u.subject, u.major_unit, u.sub_unit) for u in unit_keys if u and u.is_valid()}
        if not unit_set:
            return
        try:
            tree.itemChanged.disconnect(self._on_unit_item_changed)
        except Exception:
            pass
        try:
            self._clear_unit_selection()
            root = tree.invisibleRootItem()
            for i in range(root.childCount()):
                s_item = root.child(i)
                subject = (s_item.text(0) or "").strip()
                for j in range(s_item.childCount()):
                    m_item = s_item.child(j)
                    major = (m_item.text(0) or "").strip()
                    s_item.setExpanded(True)
                    m_item.setExpanded(True)
                    for k in range(m_item.childCount()):
                        sub_item = m_item.child(k)
                        sub = (sub_item.text(0) or "").strip()
                        if (subject, major, sub) in unit_set:
                            sub_item.setCheckState(0, Qt.Checked)
        finally:
            try:
                tree.itemChanged.connect(self._on_unit_item_changed)
            except Exception:
                pass

    def _restore_state(self, state: dict) -> None:
        """미리보기에서 돌아왔을 때 저장된 폼 상태 복원. 단원 복원 후 교재 칩/선택된 출처 갱신."""
        if not state:
            return
        try:
            # 단원
            unit_keys = state.get("unit_keys") or []
            if unit_keys:
                keys = [UnitKey(subject=k[0], major_unit=k[1], sub_unit=k[2]) for k in unit_keys if len(k) >= 3]
                self._set_unit_selection(keys)
            # 출처
            self.selected_textbook_ids = list(state.get("selected_textbook_ids") or [])
            self.selected_exam_ids = list(state.get("selected_exam_ids") or [])
            # 학년
            grade = (state.get("grade") or "").strip()
            if grade and getattr(self, "level_group", None) and getattr(self, "grade_group", None):
                level_map = {"초": "초등", "중": "중등", "고": "고등"}
                prefix = grade[0] if grade else ""
                num = grade[1:] if len(grade) > 1 else ""
                level = level_map.get(prefix, "중등")
                if getattr(self, "_level_buttons", None) and level in self._level_buttons:
                    self._level_buttons[level].setChecked(True)
                    self._update_grade_buttons(level)
                grade_btn_text = f"{num}학년" if num else ""
                for btn in self.grade_group.buttons():
                    if (btn.text() or "").strip() == grade_btn_text:
                        btn.setChecked(True)
                        break
            # 유형
            type_text = (state.get("type_text") or "").strip()
            if type_text and getattr(self, "type_group", None):
                for btn in self.type_group.buttons():
                    if (btn.text() or "").strip() == type_text:
                        btn.setChecked(True)
                        break
            # 정렬
            if getattr(self, "chk_random", None) is not None:
                self.chk_random.setChecked(bool(state.get("chk_random")))
            if getattr(self, "chk_unit_order", None) is not None:
                self.chk_unit_order.setChecked(bool(state.get("chk_unit_order", True)))
            if getattr(self, "chk_diff_order", None) is not None:
                self.chk_diff_order.setChecked(bool(state.get("chk_diff_order", True)))
            # 문항 수
            total = state.get("question_count")
            if total is not None and getattr(self, "question_count_input", None) is not None:
                v = max(1, min(9999, int(total)))
                self.question_count_input.setValue(v)
                if getattr(self, "question_slider", None) is not None:
                    self.question_slider.blockSignals(True)
                    self.question_slider.setValue(min(v, self.question_slider.maximum()))
                    self.question_slider.blockSignals(False)
            # 난이도 비율
            ratios = state.get("difficulty_ratios") or {}
            if ratios and getattr(self, "difficulty_ratio_inputs", None):
                for k, inp in self.difficulty_ratio_inputs.items():
                    if k in ratios and inp is not None:
                        inp.setText(str(ratios[k]))
            # 출처 탭(교재/기출)
            mode = state.get("source_mode") or "textbook"
            self._source_mode = mode
            if getattr(self, "btn_seg_textbook", None) is not None and getattr(self, "btn_seg_exam", None) is not None:
                self.btn_seg_textbook.setChecked(mode == "textbook")
                self.btn_seg_exam.setChecked(mode == "exam")
            # 선택된 교재/기출 표시 + 단원에 맞는 교재 칩 목록 갱신
            self.refresh_selected_sources_view()
            self._refresh_source_chips()
        except Exception:
            pass

    def init_ui(self):
        """UI 초기화 — 화이트톤 미니멀 UI (회색 박스 제거)"""
        self.setObjectName("WorksheetCreateScreen")
        self.setStyleSheet(
            """
            QWidget#WorksheetCreateScreen { background-color: #FFFFFF; }
            QWidget#WorksheetCreateScreen QFrame#ConfigCard {
                background-color: #FFFFFF; border: none; border-radius: 12px;
            }
            QWidget#WorksheetCreateScreen QTreeWidget {
                background-color: #FFFFFF; border: none; border-radius: 8px;
            }
            QWidget#WorksheetCreateScreen QLineEdit,
            QWidget#WorksheetCreateScreen QSpinBox {
                background-color: transparent; border: none;
                border-bottom: 1px solid #E0E0E0; padding: 6px 0;
                color: #000000;
            }
            QWidget#WorksheetCreateScreen QLineEdit:focus,
            QWidget#WorksheetCreateScreen QSpinBox:focus {
                border-bottom: 2px solid #007BFF;
            }
            QWidget#WorksheetCreateScreen QLabel {
                background-color: transparent; color: #222222;
            }
            QWidget#WorksheetCreateScreen QScrollArea {
                background: transparent; border: none;
            }
            """
        )

        # Root (Top-Aligned 3-Cards + Bottom Action Bar)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(30, 20, 30, 20)
        main_layout.setSpacing(15)

        # 상단: 타이틀 + 닫기(검색 없음)
        top_row = QHBoxLayout()
        title_label = QLabel("학습지 생성하기")
        title_label.setObjectName("PageTitle")
        title_label.setFont(self._font(16, bold=True))
        top_row.addWidget(title_label, alignment=Qt.AlignVCenter)
        top_row.addStretch(1)

        btn_close = QPushButton("닫기")
        btn_close.setObjectName("CloseBtn")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedHeight(38)
        btn_close.setFont(self._font(10, bold=True))
        btn_close.clicked.connect(self._on_close_clicked)
        top_row.addWidget(btn_close, alignment=Qt.AlignVCenter)
        main_layout.addLayout(top_row)

        main_layout.addSpacing(20)

        # 1) 상단 3개 카드 레이아웃 (상단 정렬 강제)
        # ✅ 카드 영역만 스크롤로 감싸서, 하단 생성 버튼이 절대 잘리지 않게 함
        cards_scroll = QScrollArea()
        cards_scroll.setWidgetResizable(True)
        cards_scroll.setFrameShape(QScrollArea.NoFrame)
        cards_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        cards_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        cards_widget = QWidget()
        cards_layout = QHBoxLayout(cards_widget)
        cards_layout.setAlignment(Qt.AlignTop)
        cards_layout.setSpacing(30)
        cards_layout.setContentsMargins(0, 0, 0, 0)

        card1 = self.create_unit_info_section()
        card2 = self.create_source_section()
        card3 = self.create_details_section()
        for c in (card1, card2, card3):
            c.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        cards_layout.addWidget(card1, 1, Qt.AlignTop)
        cards_layout.addWidget(card2, 1, Qt.AlignTop)
        cards_layout.addWidget(card3, 1, Qt.AlignTop)

        cards_scroll.setWidget(cards_widget)
        main_layout.addWidget(cards_scroll, 1)

        # 카드/버튼 사이 간격(고정값을 과하게 두면 버튼이 밀릴 수 있어 적당히)
        main_layout.addSpacing(24)

        # 2) 하단 버튼 영역 (정중앙)
        button_container = QHBoxLayout()
        button_container.addStretch(1)

        self.btn_create = QPushButton("학습지 생성")
        btn_create = self.btn_create
        btn_create.setObjectName("GenerateBtn")
        btn_create.setCursor(Qt.PointingHandCursor)
        btn_create.setFixedSize(200, 48)
        btn_create.setFont(self._font(11, bold=True))
        btn_create.clicked.connect(self.on_create_clicked)
        button_container.addWidget(btn_create, alignment=Qt.AlignCenter)

        button_container.addStretch(1)
        main_layout.addLayout(button_container)

        # 스타일은 앱 전역 테마(`ui/theme.py`)에서 관리합니다.
    
    def _font(self, size_pt: int, *, bold: bool = False) -> QFont:
        return _pick_font(size_pt, bold=bold)

    def _apply_card_shadow(self, card: QFrame) -> None:
        try:
            shadow = QGraphicsDropShadowEffect(card)
            shadow.setBlurRadius(25)
            shadow.setXOffset(0)
            shadow.setYOffset(12)
            # opacity ~0.1
            shadow.setColor(QColor(0, 0, 0, 26))
            card.setGraphicsEffect(shadow)
        except Exception:
            pass

    def _create_card(self, title: str) -> tuple[QFrame, QVBoxLayout]:
        card = QFrame()
        card.setObjectName("ConfigCard")
        self._apply_card_shadow(card)

        layout = QVBoxLayout(card)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(0)
        layout.setAlignment(Qt.AlignTop)

        lbl = QLabel(title)
        lbl.setObjectName("CardTitle")
        lbl.setFont(self._font(13, bold=True))
        lbl.setFixedHeight(32)
        lbl.setStyleSheet("padding-top: 5px; margin-bottom: 10px; font-size: 14pt; font-weight: bold;")
        layout.addWidget(lbl)
        layout.addSpacing(20)
        return card, layout

    def create_unit_info_section(self):
        """단원 정보 섹션 생성"""
        group, layout = self._create_card("1. 단원 선택")
        
        tree = QTreeWidget()
        self._unit_tree = tree
        tree.setHeaderHidden(True)
        tree.setFont(self._font(10))
        tree.setRootIsDecorated(True)  # 화살표 표시 활성화
        tree.setIndentation(25)  # 화살표와 체크박스 사이 간격 확보
        tree.setUniformRowHeights(True)
        tree.setMinimumHeight(520)
        tree.setCursor(Qt.PointingHandCursor)
        tree.setAllColumnsShowFocus(True)
        tree.setStyleSheet(
            """
            QTreeWidget {
                background-color: #FFFFFF;
                border: none;
                selection-background-color: transparent;
                selection-color: #222222;
            }
            QTreeWidget::item {
                padding: 12px 8px;
                color: #222222;
                background: transparent;
                border: none;
                outline: none;
            }
            /* 호버/선택 시 아이템 배경 (::branch는 건드리지 않아 Qt 기본 ▶/▼ 화살표가 그려지도록 함) */
            QTreeWidget::item:hover {
                background-color: #E8F0FE;
            }
            QTreeWidget::item:selected {
                background-color: #E8F0FE;
                color: #007BFF;
            }
            QTreeWidget::item:selected:focus {
                background-color: #E8F0FE;
            }
            QTreeWidget::indicator {
                width: 16px;
                height: 16px;
                border: 2px solid #94A3B8;
                border-radius: 4px;
                background-color: #ffffff;
            }
            QTreeWidget::indicator:checked {
                background-color: #2563EB;
                border-color: #2563EB;
            }
            QTreeWidget::indicator:hover {
                border-color: #2563EB;
            }
            """
        )
        pal = tree.palette()
        pal.setColor(QPalette.Highlight, Qt.transparent)
        pal.setColor(QPalette.HighlightedText, QColor(0x22, 0x22, 0x22))
        dark = QColor(0x47, 0x56, 0x69)
        pal.setColor(QPalette.Text, dark)
        pal.setColor(QPalette.WindowText, dark)
        pal.setColor(QPalette.ButtonText, dark)
        tree.setPalette(pal)

        # unit_catalog 기반 트리 구성: 과목 → 대단원 → 소단원(leaf)
        for subject, majors in UNIT_CATALOG.items():
            s_item = QTreeWidgetItem(tree)
            s_item.setText(0, subject)
            s_item.setCheckState(0, Qt.Unchecked)
            s_item.setExpanded(False)

            for major, subs in (majors or {}).items():
                m_item = QTreeWidgetItem(s_item)
                m_item.setText(0, major)
                m_item.setCheckState(0, Qt.Unchecked)
                m_item.setExpanded(False)

                for sub in (subs or []):
                    sub_item = QTreeWidgetItem(m_item)
                    sub_item.setText(0, sub)
                    sub_item.setCheckState(0, Qt.Unchecked)

        tree.itemChanged.connect(self._on_unit_item_changed)
        tree.viewport().installEventFilter(self)

        layout.addWidget(tree)
        layout.addStretch(1)
        
        return group

    def _on_unit_item_changed(self, item: QTreeWidgetItem, column: int) -> None:
        # 체크 상태를 자식으로 전파(과목/대단원 체크하면 하위 모두 체크)
        if column != 0:
            return
        if item is None:
            return

        state = item.checkState(0)
        tree = self._unit_tree
        if tree is None:
            return

        tree.blockSignals(True)
        try:
            for i in range(item.childCount()):
                child = item.child(i)
                child.setCheckState(0, state)
                # grand-children
                for j in range(child.childCount()):
                    child.child(j).setCheckState(0, state)
        finally:
            tree.blockSignals(False)

    def eventFilter(self, obj, event):
        """단원 트리: 체크박스 영역 클릭 → 체크 토글(기본 동작). 텍스트 영역 클릭 → 펼침/접기만."""
        tree = self._unit_tree
        if tree is None:
            return super().eventFilter(obj, event)
        if obj is tree.viewport() and event.type() == QEvent.MouseButtonPress and event.button() == Qt.LeftButton:
            index = tree.indexAt(event.pos())
            if index.isValid():
                item = tree.itemFromIndex(index)
                if item is not None:
                    rect = tree.visualRect(index)
                    is_text_area = event.pos().x() - rect.x() > 28
                    if item.childCount() > 0:
                        if is_text_area:
                            item.setExpanded(not item.isExpanded())
                            tree.clearSelection()
                            return True
                        # 화살표(브랜치) 영역 클릭: 트리가 펼침/접기 처리한 뒤 선택 해제
                        QTimer.singleShot(0, tree.clearSelection)
                    else:
                        # 소단원(leaf): 텍스트 영역 클릭 시 선택 효과 없음(이벤트 소비). 체크박스는 그대로 동작
                        if is_text_area:
                            tree.clearSelection()
                            return True
        return super().eventFilter(obj, event)

    def get_selected_units(self) -> List[UnitKey]:
        tree = self._unit_tree
        if tree is None:
            return []

        units: List[UnitKey] = []
        root = tree.invisibleRootItem()
        for i in range(root.childCount()):
            s_item = root.child(i)
            subject = (s_item.text(0) or "").strip()
            for j in range(s_item.childCount()):
                m_item = s_item.child(j)
                major = (m_item.text(0) or "").strip()
                for k in range(m_item.childCount()):
                    sub_item = m_item.child(k)
                    if sub_item.checkState(0) != Qt.Checked:
                        continue
                    sub = (sub_item.text(0) or "").strip()
                    uk = UnitKey(subject=subject, major_unit=major, sub_unit=sub)
                    if uk.is_valid():
                        units.append(uk)
        return units
    
    def create_source_section(self):
        """출처 섹션 생성"""
        group, layout = self._create_card("2. 출처 설정")
        # ✅ Card 2 내부 스타일만 로컬 적용 (다른 카드/네비 영향 없음)
        group.setProperty("cardRole", "source")
        group.setStyleSheet(
            """
            QLabel#SourceSectionLabel {
                color: #64748B;
                font-size: 10pt;
                font-weight: 600;
                background: transparent;
            }
            QFrame#SourceTag {
                background-color: transparent;
                border: none;
            }
            QLabel#TagIcon {
                color: #2563EB;
                padding-right: 6px;
                font-weight: 800;
                background: transparent;
            }
            QLabel#TagName {
                color: #1E293B;
                font-weight: 600;
                background: transparent;
            }
            QPushButton#DeleteBtn {
                color: #94A3B8;
                border: none;
                background: transparent;
                padding: 0px;
            }
            QPushButton#DeleteBtn:hover {
                color: #EF4444;
            }
            """
        )

        # 1) 액션 버튼 2개 (교재/기출 모달 연동)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self.btn_pick_textbook = QPushButton("교재 선택하기")
        self.btn_pick_textbook.setObjectName("SourcePickBtn")
        self.btn_pick_textbook.setCursor(Qt.PointingHandCursor)
        self.btn_pick_textbook.setMinimumHeight(44)
        self.btn_pick_textbook.setFont(self._font(11, bold=True))
        self.btn_pick_textbook.clicked.connect(self.on_select_textbooks_clicked)

        self.btn_pick_exam = QPushButton("기출 선택하기")
        self.btn_pick_exam.setObjectName("SourcePickBtn")
        self.btn_pick_exam.setCursor(Qt.PointingHandCursor)
        self.btn_pick_exam.setMinimumHeight(44)
        self.btn_pick_exam.setFont(self._font(11, bold=True))
        self.btn_pick_exam.clicked.connect(self.on_select_exams_clicked)

        btn_row.addWidget(self.btn_pick_textbook, 1)
        btn_row.addWidget(self.btn_pick_exam, 1)
        layout.addLayout(btn_row)

        layout.addSpacing(20)

        # 2) 선택된 교재
        tb_title = QLabel("선택된 교재")
        tb_title.setObjectName("SourceSectionLabel")
        tb_title.setFont(self._font(10, bold=False))
        layout.addWidget(tb_title)
        layout.addSpacing(10)

        self.selected_tb_scroll = QScrollArea()
        self.selected_tb_scroll.setObjectName("SelectedContainer")
        self.selected_tb_scroll.setWidgetResizable(True)
        self.selected_tb_scroll.setFrameShape(QScrollArea.NoFrame)
        self.selected_tb_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.selected_tb_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.selected_tb_scroll.setFixedHeight(160)

        tb_widget = QWidget()
        tb_widget.setObjectName("SelectedContainerWidget")
        tb_layout = QVBoxLayout(tb_widget)
        tb_layout.setContentsMargins(12, 12, 12, 12)
        tb_layout.setSpacing(12)

        self._tb_empty_hint = QLabel("선택된 교재가 없습니다")
        self._tb_empty_hint.setObjectName("EmptyHint")
        self._tb_empty_hint.setFont(self._font(10, bold=True))
        self._tb_empty_hint.setAlignment(Qt.AlignCenter)
        tb_layout.addWidget(self._tb_empty_hint, 1)

        self._tb_tags_wrap = QWidget()
        self._tb_tags_wrap.setObjectName("SelectedTagsWrap")
        self._selected_tb_flow = FlowLayout(self._tb_tags_wrap, spacing=10)
        self._tb_tags_wrap.setLayout(self._selected_tb_flow)
        tb_layout.addWidget(self._tb_tags_wrap, 0)

        self.selected_tb_scroll.setWidget(tb_widget)
        layout.addWidget(self.selected_tb_scroll)

        layout.addSpacing(20)

        # 3) 선택된 내신기출
        ex_title = QLabel("선택된 내신기출")
        ex_title.setObjectName("SourceSectionLabel")
        ex_title.setFont(self._font(10, bold=False))
        layout.addWidget(ex_title)
        layout.addSpacing(10)

        self.selected_ex_scroll = QScrollArea()
        self.selected_ex_scroll.setObjectName("SelectedContainer")
        self.selected_ex_scroll.setWidgetResizable(True)
        self.selected_ex_scroll.setFrameShape(QScrollArea.NoFrame)
        self.selected_ex_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.selected_ex_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.selected_ex_scroll.setFixedHeight(160)

        ex_widget = QWidget()
        ex_widget.setObjectName("SelectedContainerWidget")
        ex_layout = QVBoxLayout(ex_widget)
        ex_layout.setContentsMargins(12, 12, 12, 12)
        ex_layout.setSpacing(12)

        self._ex_empty_hint = QLabel("선택된 내신기출이 없습니다")
        self._ex_empty_hint.setObjectName("EmptyHint")
        self._ex_empty_hint.setFont(self._font(10, bold=True))
        self._ex_empty_hint.setAlignment(Qt.AlignCenter)
        ex_layout.addWidget(self._ex_empty_hint, 1)

        self._ex_tags_wrap = QWidget()
        self._ex_tags_wrap.setObjectName("SelectedTagsWrap")
        self._selected_ex_flow = FlowLayout(self._ex_tags_wrap, spacing=10)
        self._ex_tags_wrap.setLayout(self._selected_ex_flow)
        ex_layout.addWidget(self._ex_tags_wrap, 0)

        self.selected_ex_scroll.setWidget(ex_widget)
        layout.addWidget(self.selected_ex_scroll)

        layout.addStretch(1)

        self.refresh_selected_sources_view()
        return group

    def _on_source_mode_changed(self, btn: QPushButton) -> None:
        try:
            self._source_mode = "exam" if btn is self.btn_seg_exam else "textbook"
        except Exception:
            self._source_mode = "textbook"
        self._safe_refresh_sources()

    def _safe_refresh_sources(self) -> None:
        """
        출처 UI 갱신 중 예외가 나도 앱이 종료되지 않도록 보호.
        (특히 Mongo 연결 불안정/데이터 이상 시 크래시 방지)
        """
        try:
            self._refresh_source_chips()
            self.refresh_selected_sources_view()
        except Exception as e:
            # 콘솔/로그가 없는 실행 환경에서도 원인 확인 가능하게 메시지 출력
            try:
                QMessageBox.critical(
                    self,
                    "출처 로딩 오류",
                    "기출/교재 목록을 불러오는 중 오류가 발생했습니다.\n"
                    "DB 연결 상태 또는 데이터 형식을 확인해주세요.\n\n"
                    f"{type(e).__name__}: {e}",
                )
            except Exception:
                pass
            try:
                print("=== Source refresh crash ===")
                traceback.print_exc()
            except Exception:
                pass

    def _clear_flow(self, flow: Optional[FlowLayout]) -> None:
        if flow is None:
            return
        while flow.count():
            it = flow.takeAt(0)
            if it is None:
                continue
            w = it.widget()
            if w is not None:
                w.setParent(None)

    def _make_source_chip(self, text: str, *, on_click: Callable[[], None], disabled: bool = False) -> QPushButton:
        b = QPushButton(text)
        b.setObjectName("SourceChip")
        b.setCursor(Qt.PointingHandCursor)
        b.setCheckable(False)
        b.setEnabled(not disabled)
        b.setFont(self._font(9, bold=True))
        b.clicked.connect(on_click)
        return b

    def _make_selected_tag(self, icon: str, text: str, *, on_remove: Callable[[], None]) -> QFrame:
        tag = QFrame()
        tag.setObjectName("SourceTag")
        hl = QHBoxLayout(tag)
        hl.setContentsMargins(10, 6, 10, 6)
        hl.setSpacing(0)

        ico = QLabel(icon)
        ico.setObjectName("TagIcon")
        ico.setFont(self._font(10, bold=True))
        hl.addWidget(ico, alignment=Qt.AlignVCenter)

        lbl = QLabel(text)
        lbl.setObjectName("TagName")
        lbl.setFont(self._font(9, bold=True))
        hl.addWidget(lbl, alignment=Qt.AlignVCenter)

        hl.addSpacing(8)
        hl.addStretch(1)

        btn = QPushButton("×")
        btn.setObjectName("DeleteBtn")
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFixedSize(16, 16)
        btn.setFont(self._font(10, bold=True))
        # PyQt5 clicked(bool) 인자가 on_remove에 전달되면 tid_/eid_로 오인되어 삭제가 동작하지 않으므로 람다로 무시
        btn.clicked.connect(lambda: on_remove())
        hl.addWidget(btn, alignment=Qt.AlignVCenter)
        return tag

    def _eligible_textbooks(self) -> List[Tuple[str, str]]:
        """
        Returns: [(id, label), ...]  (단원 선택과 일치하는 교재만)
        """
        units = self.get_selected_units()
        if not units:
            return []
        unit_set = {(u.subject, u.major_unit, u.sub_unit) for u in units if u and u.is_valid()}
        out: List[Tuple[str, str]] = []
        for t in self.textbook_repo.list_all():
            if not t or not t.id:
                continue
            key = ((t.subject or "").strip(), (t.major_unit or "").strip(), (t.sub_unit or "").strip())
            if key not in unit_set:
                continue
            out.append((str(t.id), (t.name or "").strip() or str(t.id)))
        return out

    def _eligible_exams(self) -> List[Tuple[str, str]]:
        out: List[Tuple[str, str]] = []
        for e in self.exam_repo.list_all():
            if not e or not e.id:
                continue
            label = f"{e.school_name} {e.grade} {e.semester} {e.exam_type} ({e.year})"
            out.append((str(e.id), label.strip() or str(e.id)))
        return out

    def _refresh_source_chips(self) -> None:
        flow = self._result_flow
        if flow is None:
            return
        self._clear_flow(flow)

        q = ""
        if self._source_search is not None:
            q = (self._source_search.text() or "").strip().lower()

        mode = self._source_mode or "textbook"
        if mode == "textbook":
            # 단원 미선택이면 교재 칩을 비움
            items = self._eligible_textbooks()
            selected = set(str(x) for x in (self.selected_textbook_ids or []))
        else:
            items = self._eligible_exams()
            selected = set(str(x) for x in (self.selected_exam_ids or []))

        # 검색 필터 + 선택된 항목 숨김
        filtered: List[Tuple[str, str]] = []
        for _id, label in items:
            if str(_id) in selected:
                continue
            if q:
                if q not in (label or "").lower():
                    continue
            filtered.append((_id, label))

        # 결과가 너무 많으면 상위 N개만(UX 보호)
        filtered = filtered[:40]

        if not filtered:
            msg = "단원을 선택하면 교재가 표시됩니다." if mode == "textbook" and not self.get_selected_units() else "검색 결과가 없습니다."
            lbl = QLabel(msg)
            lbl.setFont(self._font(9))
            lbl.setStyleSheet("color:#94A3B8;")
            flow.addWidget(lbl)
            return

        for _id, label in filtered:
            def _add(id_=_id):
                if mode == "textbook":
                    if str(id_) not in set(str(x) for x in self.selected_textbook_ids):
                        self.selected_textbook_ids.append(str(id_))
                else:
                    if str(id_) not in set(str(x) for x in self.selected_exam_ids):
                        self.selected_exam_ids.append(str(id_))
                self.refresh_selected_sources_view()
                self._refresh_source_chips()

            chip = self._make_source_chip(label, on_click=_add, disabled=False)
            chip.setToolTip(label)
            flow.addWidget(chip)

    def refresh_selected_sources_view(self) -> None:
        self._clear_flow(self._selected_tb_flow)
        self._clear_flow(self._selected_ex_flow)

        # id -> label
        tb_map = {str(t.id): t for t in (self.textbook_repo.list_all() or []) if t and t.id}
        ex_map = {str(e.id): e for e in (self.exam_repo.list_all() or []) if e and e.id}

        # 빈 상태 안내
        try:
            if getattr(self, "_tb_empty_hint", None) is not None:
                self._tb_empty_hint.setVisible(not bool(self.selected_textbook_ids))
            if getattr(self, "_tb_tags_wrap", None) is not None:
                self._tb_tags_wrap.setVisible(bool(self.selected_textbook_ids))
        except Exception:
            pass
        try:
            if getattr(self, "_ex_empty_hint", None) is not None:
                self._ex_empty_hint.setVisible(not bool(self.selected_exam_ids))
            if getattr(self, "_ex_tags_wrap", None) is not None:
                self._ex_tags_wrap.setVisible(bool(self.selected_exam_ids))
        except Exception:
            pass

        if self._selected_tb_flow is not None:
            for tid in list(self.selected_textbook_ids or []):
                t = tb_map.get(str(tid))
                name = (t.name if t else str(tid)) or str(tid)

                def _rm(tid_=str(tid)):
                    self.selected_textbook_ids = [x for x in self.selected_textbook_ids if str(x) != tid_]
                    self.refresh_selected_sources_view()

                self._selected_tb_flow.addWidget(self._make_selected_tag("📚", name, on_remove=_rm))

        if self._selected_ex_flow is not None:
            for eid in list(self.selected_exam_ids or []):
                e = ex_map.get(str(eid))
                label = f"{e.school_name} {e.grade} {e.semester} {e.exam_type} ({e.year})" if e else str(eid)

                def _rm(eid_=str(eid)):
                    self.selected_exam_ids = [x for x in self.selected_exam_ids if str(x) != eid_]
                    self.refresh_selected_sources_view()

                self._selected_ex_flow.addWidget(self._make_selected_tag("📝", label, on_remove=_rm))

    def on_select_textbooks_clicked(self) -> None:
        units = self.get_selected_units()
        if not units:
            QMessageBox.information(self, "단원 선택 필요", "교재 선택 전에 단원을 먼저 선택해 주세요.")
            return
        try:
            dlg = TextbookMultiSelectDialog(
                self.db_connection,
                units=units,
                preselected_ids=self.selected_textbook_ids,
                parent=self,
            )
            if dlg.exec_() == dlg.Accepted:
                self.selected_textbook_ids = dlg.selected_ids()
                self.refresh_selected_sources_view()
        except Exception as e:
            QMessageBox.critical(self, "교재 선택 오류", str(e))

    def on_select_exams_clicked(self) -> None:
        try:
            dlg = ExamMultiSelectDialog(
                self.db_connection,
                preselected_ids=self.selected_exam_ids,
                parent=self,
            )
            if dlg.exec_() == dlg.Accepted:
                self.selected_exam_ids = dlg.selected_ids()
                self.refresh_selected_sources_view()
        except Exception as e:
            QMessageBox.critical(self, "기출 선택 오류", str(e))
    
    def create_details_section(self):
        """세부 옵션 섹션 — 좌측 카드와 동일한 카드 형태, [항목 이름 - 설정 요소] 리스트형 배치"""
        group, layout = self._create_card("3. 세부 옵션")
        group.setProperty("cardRole", "details")
        layout.setContentsMargins(20, 14, 20, 14)
        layout.setSpacing(6)
        group.setStyleSheet(
            """
            QFrame[cardRole="details"] QLineEdit, QFrame[cardRole="details"] QLineEdit#MiniInput {
                background: transparent; border: none; border-bottom: 1px solid #EEEEEE;
                padding: 4px 0; color: #000000; min-height: 26px;
            }
            QFrame[cardRole="details"] QSpinBox {
                background: transparent; border: none; border-bottom: 1px solid #EEEEEE;
                padding: 4px 0; color: #000000; min-height: 26px;
            }
            QFrame[cardRole="details"] QPushButton#FilterChip {
                background: transparent; border: none; color: #777777;
                padding: 4px 10px; font-size: 10pt; border-radius: 12px;
                min-height: 26px;
            }
            QFrame[cardRole="details"] QPushButton#FilterChip:hover {
                background-color: #F5F5F5;
            }
            QFrame[cardRole="details"] QPushButton#FilterChip:checked {
                background-color: #E8F0FE; color: #007BFF; font-weight: bold;
                border-radius: 12px;
            }
            QFrame[cardRole="details"] QLabel#SectionLabel {
                color: #333333; font-weight: bold; font-size: 10pt;
                background: transparent;
            }
            """
        )
        try:
            group.style().unpolish(group)
            group.style().polish(group)
        except Exception:
            pass

        def _opt_label(text: str) -> QLabel:
            lbl = QLabel(text)
            lbl.setObjectName("SectionLabel")
            lbl.setFont(self._font(10, bold=True))
            return lbl

        def _divider() -> QFrame:
            line = QFrame()
            line.setObjectName("DividerLine")
            line.setFixedHeight(1)
            line.setStyleSheet("background-color: #F5F5F5; border: none;")
            return line

        def _option_block(title: str, content: QWidget):
            """라벨 상단·콘텐츠 하단 블록. 컴팩트 간격으로 한 화면 배치."""
            layout.addWidget(_opt_label(title))
            layout.addSpacing(4)
            layout.addWidget(content)
            layout.addWidget(_divider())
            layout.addSpacing(6)

        # 1. 난이도 비율 (컴팩트)
        diff_row = QWidget()
        diff_layout = QHBoxLayout(diff_row)
        diff_layout.setContentsMargins(0, 0, 0, 0)
        diff_layout.setSpacing(8)
        difficulty_options = [("킬", "최상"), ("상", "상"), ("중", "중"), ("하", "하")]
        self.difficulty_ratio_inputs = {}
        for key, label_text in difficulty_options:
            v_box = QVBoxLayout()
            v_box.setSpacing(2)
            l = QLabel(label_text)
            l.setFont(self._font(9, bold=True))
            l.setAlignment(Qt.AlignCenter)
            v_box.addWidget(l)
            inp = QLineEdit()
            inp.setObjectName("MiniInput")
            inp.setPlaceholderText("0")
            inp.setFixedWidth(48)
            inp.setFixedHeight(26)
            inp.setFont(self._font(9))
            inp.setAlignment(Qt.AlignCenter)
            inp.setValidator(QIntValidator(0, 100, self))
            self.difficulty_ratio_inputs[key] = inp
            v_box.addWidget(inp, alignment=Qt.AlignCenter)
            diff_layout.addLayout(v_box)
        diff_layout.addStretch(1)
        self.diff_sum_label = QLabel("합계 100%")
        self.diff_sum_label.setFont(self._font(9, bold=True))
        self.diff_sum_label.setObjectName("SectionHint")
        diff_layout.addWidget(self.diff_sum_label, alignment=Qt.AlignVCenter)

        _option_block("난이도 비율(%)", diff_row)

        self.difficulty_ratio_inputs["킬"].setText("30")
        self.difficulty_ratio_inputs["상"].setText("20")
        self.difficulty_ratio_inputs["중"].setText("20")
        self.difficulty_ratio_inputs["하"].setText("30")

        def _update_diff_sum() -> None:
            total = 0
            ok = True
            for k in ["킬", "상", "중", "하"]:
                s = (self.difficulty_ratio_inputs.get(k).text() if self.difficulty_ratio_inputs.get(k) else "").strip()
                if s == "":
                    ok = False
                    continue
                try:
                    total += int(s)
                except Exception:
                    ok = False
            if ok and total == 100:
                self.diff_sum_label.setText("합계 100%")
                self.diff_sum_label.setStyleSheet("color: #16A34A;")
            else:
                self.diff_sum_label.setText(f"합계 {total}%")
                self.diff_sum_label.setStyleSheet("color: #DC2626;")

        for _k, _inp in self.difficulty_ratio_inputs.items():
            try:
                _inp.textChanged.connect(_update_diff_sum)
            except Exception:
                pass
        _update_diff_sum()

        # 2. 문항 수 (컴팩트)
        count_inline = QFrame()
        count_inline.setObjectName("CountInline")
        count_row = QHBoxLayout(count_inline)
        count_row.setContentsMargins(0, 0, 0, 0)
        count_row.setSpacing(8)
        self.question_slider = QSlider(Qt.Horizontal)
        self.question_slider.setRange(1, 500)
        self.question_slider.setValue(50)
        self.question_slider.setFixedHeight(26)
        self.question_count_input = QSpinBox()
        self.question_count_input.setRange(1, 9999)
        self.question_count_input.setValue(50)
        self.question_count_input.setFixedWidth(72)
        self.question_count_input.setFixedHeight(26)
        self.question_count_input.setSuffix("")
        self.question_count_input.setFont(self._font(9, bold=True))
        try:
            self.question_count_input.setAlignment(Qt.AlignCenter)
        except Exception:
            pass

        def _sync_from_slider(v: int) -> None:
            try:
                self.question_count_input.blockSignals(True)
                self.question_count_input.setValue(int(v))
            finally:
                self.question_count_input.blockSignals(False)

        def _sync_from_spin(v: int) -> None:
            try:
                vmax = int(self.question_slider.maximum())
                if int(v) <= vmax:
                    self.question_slider.blockSignals(True)
                    self.question_slider.setValue(int(v))
            finally:
                self.question_slider.blockSignals(False)

        self.question_slider.valueChanged.connect(_sync_from_slider)
        self.question_count_input.valueChanged.connect(_sync_from_spin)
        count_row.addWidget(self.question_slider, 1)
        count_row.addWidget(self.question_count_input, 0)

        _option_block("문항 수", count_inline)

        # 3. 학년 — 2단: [초등/중등/고등] → 해당 학년 버튼 (컴팩트)
        def _mk_chip(text: str) -> QPushButton:
            b = QPushButton(text)
            b.setObjectName("FilterChip")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setFont(self._font(9, bold=True))
            return b

        grade_wrap = QWidget()
        gv = QVBoxLayout(grade_wrap)
        gv.setContentsMargins(0, 0, 0, 0)
        gv.setSpacing(6)
        self.level_group = QButtonGroup(self)
        self.level_group.setExclusive(True)
        level_row = QHBoxLayout()
        level_row.setContentsMargins(0, 0, 0, 0)
        level_row.setSpacing(6)
        self._level_buttons = {}
        for level in ["초등", "중등", "고등"]:
            btn = _mk_chip(level)
            self._level_buttons[level] = btn
            self.level_group.addButton(btn)
            btn.clicked.connect(lambda checked, l=level: self._update_grade_buttons(l))
            level_row.addWidget(btn)
        level_row.addStretch(1)
        gv.addLayout(level_row)
        self._grade_container = QWidget()
        self._grade_layout = QHBoxLayout(self._grade_container)
        self._grade_layout.setContentsMargins(0, 0, 0, 0)
        self._grade_layout.setSpacing(6)
        self.grade_group = QButtonGroup(self)
        self.grade_group.setExclusive(True)
        gv.addWidget(self._grade_container)
        self._level_buttons["중등"].setChecked(True)
        self._update_grade_buttons("중등")
        _option_block("학년", grade_wrap)

        # 4. 유형 (컴팩트 칩)
        type_wrap = QWidget()
        type_layout = QHBoxLayout(type_wrap)
        type_layout.setContentsMargins(0, 0, 0, 0)
        type_layout.setSpacing(6)
        self.type_group = QButtonGroup(self)
        self.type_group.setExclusive(True)
        for t in ["TEST", "과제", "교재"]:
            b = _mk_chip(t)
            if t == "TEST":
                b.setChecked(True)
            self.type_group.addButton(b)
            type_layout.addWidget(b)
        type_layout.addStretch(1)
        _option_block("유형", type_wrap)

        # 5. 정렬 (컴팩트 칩)
        order_wrap = QWidget()
        order_layout = QHBoxLayout(order_wrap)
        order_layout.setContentsMargins(0, 0, 0, 0)
        order_layout.setSpacing(6)
        self.chk_random = _mk_chip("랜덤")
        self.chk_unit_order = _mk_chip("단원 순서")
        self.chk_diff_order = _mk_chip("난이도 순서")
        self.chk_random.setCheckable(True)
        self.chk_unit_order.setCheckable(True)
        self.chk_diff_order.setCheckable(True)
        self.chk_unit_order.setChecked(True)
        self.chk_diff_order.setChecked(True)
        self.chk_random.toggled.connect(self._on_random_changed)
        order_layout.addWidget(self.chk_random)
        order_layout.addWidget(self.chk_unit_order)
        order_layout.addWidget(self.chk_diff_order)
        order_layout.addStretch(1)
        layout.addWidget(_opt_label("정렬"))
        layout.addSpacing(6)
        layout.addWidget(order_wrap)
        layout.addWidget(_divider())

        return group

    def _on_random_changed(self, checked) -> None:
        is_on = bool(checked)
        # 랜덤이면 다른 정렬 체크 불가(요구사항)
        self.chk_unit_order.setEnabled(not is_on)
        self.chk_diff_order.setEnabled(not is_on)
        if is_on:
            self.chk_unit_order.setChecked(False)
            self.chk_diff_order.setChecked(False)

    def _read_ratios(self) -> Optional[dict]:
        ratios = {}
        total = 0
        for k in ["킬", "상", "중", "하"]:
            s = (self.difficulty_ratio_inputs.get(k).text() if self.difficulty_ratio_inputs.get(k) else "").strip()
            if s == "":
                return None
            try:
                v = int(s)
            except Exception:
                return None
            ratios[k] = v
            total += v
        if total != 100:
            return None
        return ratios

    def _update_grade_buttons(self, level: str) -> None:
        """학교급 선택에 따라 학년 버튼 동적 교체"""
        # 기존 학년 버튼 제거
        for btn in self.grade_group.buttons():
            self.grade_group.removeButton(btn)
            btn.setParent(None)
        while self._grade_layout.count():
            item = self._grade_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        level_prefix = {"초등": "초", "중등": "중", "고등": "고"}.get(level, "")
        if level == "초등":
            grades = [f"{i}학년" for i in range(1, 7)]
        else:
            grades = ["1학년", "2학년", "3학년"]

        def _mk_chip(text: str) -> QPushButton:
            b = QPushButton(text)
            b.setObjectName("FilterChip")
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setFont(self._font(9, bold=True))
            return b

        for g in grades:
            btn = _mk_chip(g)
            self.grade_group.addButton(btn)
            self._grade_layout.addWidget(btn)
        self._grade_layout.addStretch(1)
        if grades:
            self.grade_group.buttons()[0].setChecked(True)

    def _selected_grade(self) -> str:
        """학교급 + 학년 조합으로 기존 형식(초1, 중2, 고3 등) 반환"""
        try:
            level_btn = self.level_group.checkedButton() if getattr(self, "level_group", None) else None
            grade_btn = self.grade_group.checkedButton() if getattr(self, "grade_group", None) else None
            if not level_btn or not grade_btn:
                return ""
            level_text = (level_btn.text() or "").strip()
            grade_text = (grade_btn.text() or "").strip()
            prefix = {"초등": "초", "중등": "중", "고등": "고"}.get(level_text, "")
            num = grade_text.replace("학년", "").strip() if grade_text else ""
            if prefix and num:
                return f"{prefix}{num}"
            return ""
        except Exception:
            return ""

    def _selected_type(self) -> str:
        try:
            btn = self.type_group.checkedButton()
            return (btn.text() if btn else "").strip()
        except Exception:
            return ""

    def on_create_clicked(self) -> None:
        # 1) 입력 수집/검증
        units = self.get_selected_units()
        if not units:
            QMessageBox.warning(self, "입력 오류", "단원을 1개 이상 선택해 주세요(소단원까지).")
            return

        if not (self.selected_textbook_ids or self.selected_exam_ids):
            QMessageBox.warning(self, "입력 오류", "출처(교재 또는 내신기출)를 1개 이상 선택해 주세요.")
            return

        # ✅ 문항 수는 직접 입력(SpinBox)을 기준으로 함 (무제한 입력 보장)
        total = int(getattr(self, "question_count_input", None).value() if getattr(self, "question_count_input", None) else self.question_slider.value() or 0)
        if total <= 0:
            QMessageBox.warning(self, "입력 오류", "문항수는 1 이상이어야 합니다.")
            return

        ratios = self._read_ratios()
        if not ratios:
            QMessageBox.warning(self, "입력 오류", "난이도 비율(최상/상/중/하)의 합계를 100으로 맞춰 주세요.")
            return

        if self.chk_random.isChecked():
            order = OrderOptions(randomize=True, order_by_unit=False, order_by_difficulty=False)
        else:
            order = OrderOptions(
                randomize=False,
                order_by_unit=self.chk_unit_order.isChecked(),
                order_by_difficulty=self.chk_diff_order.isChecked(),
            )

        sources = SelectedSources(
            textbook_ids=list(self.selected_textbook_ids),
            exam_ids=list(self.selected_exam_ids),
        )

        # 2) 선택 엔진 실행
        try:
            result = self.worksheet_service.select_problems(
                units=units,
                sources=sources,
                total_count=total,
                difficulty_ratios=ratios,
                order=order,
                seed=None,  # 완전 랜덤
            )
        except Exception as e:
            QMessageBox.critical(self, "실행 실패", str(e))
            return

        self._last_selected_problem_ids = list(result.selected_problem_ids)

        # 미리보기에서 돌아올 때 복원할 폼 상태 저장
        units = self.get_selected_units()
        self._saved_state_for_restore = {
            "unit_keys": [(u.subject, u.major_unit, u.sub_unit) for u in units if u and u.is_valid()],
            "selected_textbook_ids": list(self.selected_textbook_ids),
            "selected_exam_ids": list(self.selected_exam_ids),
            "grade": self._selected_grade(),
            "type_text": self._selected_type(),
            "chk_random": self.chk_random.isChecked(),
            "chk_unit_order": self.chk_unit_order.isChecked(),
            "chk_diff_order": self.chk_diff_order.isChecked(),
            "question_count": int(getattr(self, "question_count_input", None).value() if getattr(self, "question_count_input", None) else self.question_slider.value() or 0),
            "difficulty_ratios": self._read_ratios() or {},
            "source_mode": self._source_mode or "textbook",
        }

        # 3) 문항 편집 화면으로 이동(요구사항: 먼저 배치/미리보기/드래그 편집)
        payload = {
            "draft": {
                # Step 2에서 입력 (초기값은 비움)
                "title": "",
                "creator": "",
                "grade": self._selected_grade(),
                # 목록/뱃지 표준: 출처 기반으로 자동 결정
                # - 내신기출: exam만
                # - 시중교재: textbook만
                # - 통합: 둘 다
                "type_text": (
                    "통합"
                    if (self.selected_textbook_ids and self.selected_exam_ids)
                    else ("내신기출" if self.selected_exam_ids else "시중교재")
                ),
                # Step 2에서 설정 (초기값은 False)
                "option_unit_tag": False,
                "option_source_tag": False,
                "option_difficulty_tag": False,
                "requested_total": int(result.requested_total),
                "actual_total": int(result.actual_total),
                "warnings": list(result.warnings or []),
                "problem_ids": list(result.selected_problem_ids),
            }
        }
        self.preview_requested.emit(payload)

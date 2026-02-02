"""
- 고정 폭 180px / 배경 #FFFFFF / 우측 1px 구분선 #E2E8F0
- 메뉴 아이템 높이 44px, 외부 여백 margin(2px 10px 느낌)
- 활성화: 배경 없이 텍스트/아이콘만 #2563EB + Bold (포인트 라인 제거)
- hover: #F9FAFB 배경 + 진한 텍스트
- 포커스 점선 제거: 전부 Qt.NoFocus
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional, Tuple

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QVBoxLayout,
    QWidget,
)

try:
    import qtawesome as qta  # type: ignore

    _QTA_AVAILABLE = True
except Exception:
    qta = None
    _QTA_AVAILABLE = False


@dataclass(frozen=True)
class _MenuItem:
    icon: str
    text: str


def _qta_pixmap(icon_name: str, color: str, size: int) -> Optional[QPixmap]:
    if not _QTA_AVAILABLE or qta is None:
        return None
    try:
        return qta.icon(icon_name, color=color).pixmap(size, size)
    except Exception:
        return None


class _SidebarMenuButton(QFrame):
    clicked = pyqtSignal(str)

    def __init__(self, key: str, icon_name: str, label: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._key = key
        self._icon_name = icon_name
        self._label_text = label
        self._is_hovered = False
        self._is_active = False

        self.setFocusPolicy(Qt.NoFocus)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(44)
        self.setObjectName("sidebarMenuButton")

        self._root = QHBoxLayout(self)
        self._root.setContentsMargins(0, 0, 0, 0)
        self._root.setSpacing(0)

        # 내부 컨텐츠 영역(패딩-left 15px, 아이콘-텍스트 spacing 15px)
        content = QWidget()
        content.setFocusPolicy(Qt.NoFocus)
        content_layout = QHBoxLayout(content)
        content_layout.setContentsMargins(15, 0, 12, 0)
        content_layout.setSpacing(15)

        self._icon = QLabel()
        self._icon.setFixedSize(18, 18)
        self._icon.setAlignment(Qt.AlignCenter)
        self._icon.setFocusPolicy(Qt.NoFocus)

        self._text = QLabel(label)
        # Header 메뉴와 동일 톤(폰트/크기/굵기)
        f = QFont("Inter")
        if not f.exactMatch():
            f = QFont("Pretendard")
        if not f.exactMatch():
            f = QFont("맑은 고딕")
        f.setPointSize(12)
        f.setWeight(QFont.DemiBold)  # 600
        self._text.setFont(f)
        self._text.setFocusPolicy(Qt.NoFocus)

        content_layout.addWidget(self._icon, alignment=Qt.AlignVCenter)
        content_layout.addWidget(self._text, alignment=Qt.AlignVCenter)
        content_layout.addStretch(1)

        self._root.addWidget(content)

        self._apply_visual()

    def set_active(self, active: bool) -> None:
        self._is_active = bool(active)
        if self._is_active:
            self._is_hovered = False
        self._apply_visual()

    def enterEvent(self, event):  # noqa: N802 (Qt naming)
        if not self._is_active:
            self._is_hovered = True
            self._apply_visual()
        super().enterEvent(event)

    def leaveEvent(self, event):  # noqa: N802 (Qt naming)
        if not self._is_active:
            self._is_hovered = False
            self._apply_visual()
        super().leaveEvent(event)

    def mousePressEvent(self, event):  # noqa: N802 (Qt naming)
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self._key)
        super().mousePressEvent(event)

    def _colors(self) -> Tuple[str, str, str]:
        """
        Returns (bg, text_color, icon_color)
        """
        if self._is_active:
            return "transparent", "#2563EB", "#2563EB"
        if self._is_hovered:
            return "#F9FAFB", "#111827", "#111827"
        return "transparent", "#6B7280", "#6B7280"

    def _apply_visual(self) -> None:
        bg, text_color, icon_color = self._colors()

        # 배경/라운드/활성 라인
        self.setStyleSheet(
            f"""
            QFrame#sidebarMenuButton {{
                background-color: {bg};
                border: none;
                border-radius: 10px;
            }}
            QFrame#sidebarMenuButton:hover {{
                border: none;
            }}
            QLabel {{
                border: none;
                background: transparent;
                outline: none;
            }}
            """
        )

        self._text.setStyleSheet(
            f"color: {text_color}; border: none; background: transparent; outline: none;"
        )
        # 활성 상태는 더 굵게(헤더 active 톤과 맞춤)
        f = self._text.font()
        f.setWeight(QFont.ExtraBold if self._is_active else QFont.DemiBold)
        self._text.setFont(f)

        pm = _qta_pixmap(self._icon_name, icon_color, 18)
        if pm is not None:
            self._icon.setPixmap(pm)
            self._icon.setText("")
        else:
            # qtawesome 미설치 환경 fallback: 불필요한 bullet 제거(빈 자리 유지)
            self._icon.setText("")
            self._icon.setStyleSheet("border: none; background: transparent;")


class Sidebar(QFrame):
    """사이드바 위젯 (MainWindow와 호환 API 유지)"""

    menu_clicked = pyqtSignal(str)

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(180)
        self.setFocusPolicy(Qt.NoFocus)

        self.current_menu: Optional[str] = None
        self.menu_buttons: Dict[str, _SidebarMenuButton] = {}

        self._init_ui()

    def _init_ui(self) -> None:
        self.setStyleSheet(
            """
            QFrame#sidebar {
                background-color: #FFFFFF;
                border-right: 1px solid #F1F5F9;
            }
            /* 충돌 방지: 사이드바 내부 스타일 초기화(override) */
            QFrame#sidebar * {
                border: none;
                background: transparent;
                outline: none;
            }
            """
        )

        root = QVBoxLayout(self)
        # 헤더 바로 아래에서 자연스럽게 시작하도록 상단 여백만 살짝 줄임
        root.setContentsMargins(0, 12, 0, 16)
        root.setSpacing(0)

        menus = [
            _MenuItem("fa5s.file-alt", "학습지"),
            _MenuItem("fa5s.book", "교재"),
            _MenuItem("fa5s.database", "교재DB"),
            _MenuItem("fa5s.history", "기출DB"),
            _MenuItem("fa5s.trash", "휴지통"),
        ]

        for item in menus:
            btn = _SidebarMenuButton(item.text, item.icon, item.text)
            btn.clicked.connect(self.on_menu_clicked)

            # 외부 여백: margin 2px 10px 느낌을 wrapper로 구현
            wrapper = QWidget()
            wrapper.setFocusPolicy(Qt.NoFocus)
            w_layout = QHBoxLayout(wrapper)
            w_layout.setContentsMargins(10, 2, 10, 2)
            w_layout.setSpacing(0)
            w_layout.addWidget(btn)

            self.menu_buttons[item.text] = btn
            root.addWidget(wrapper)

        root.addStretch(1)

        # 초기 활성화는 선택하지 않음(기존 동작 유지)
        self.current_menu = None

    def clear_selection(self) -> None:
        for btn in self.menu_buttons.values():
            btn.set_active(False)
        self.current_menu = None

    def on_menu_clicked(self, menu_name: str) -> None:
        for name, btn in self.menu_buttons.items():
            btn.set_active(name == menu_name)
        self.current_menu = menu_name
        self.menu_clicked.emit(menu_name)


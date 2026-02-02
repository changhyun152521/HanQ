"""
상단 헤더 컴포넌트 (Seamless, Overwrite from scratch)

- 높이 70px
- 로고: 아이콘 + 텍스트(하나의 유닛)
- 메뉴: [수업준비, 수업, 관리] (검색창 제거)
- 우측: 알림 + 유저 텍스트 + 로그아웃
"""

from __future__ import annotations

from typing import Dict, Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QPixmap, QColor
from PyQt5.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QWidget,
)

try:
    import qtawesome as qta  # type: ignore

    _QTA_AVAILABLE = True
except Exception:
    qta = None
    _QTA_AVAILABLE = False


def _qta_pixmap(icon_name: str, color: str, size: int) -> Optional[QPixmap]:
    if not _QTA_AVAILABLE or qta is None:
        return None
    try:
        return qta.icon(icon_name, color=color).pixmap(size, size)
    except Exception:
        return None


def _fallback_logo_pixmap(color_hex: str = "#2563EB", size: int = 24) -> QPixmap:
    """qtawesome 미설치/실패 시에도 항상 보이는 로고 아이콘."""
    from PyQt5.QtCore import QRectF
    from PyQt5.QtGui import QPainter, QPen

    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)

    c = QColor(color_hex)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(c)
    pen.setWidthF(1.8)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)

    # 간단한 책 모양(좌/우 페이지)
    pad = 4.0
    mid = size / 2.0
    top = 5.0
    bottom = size - 5.0
    w = mid - pad
    h = bottom - top
    left = QRectF(pad, top, w, h)
    right = QRectF(mid, top, w, h)
    painter.drawRoundedRect(left, 3.5, 3.5)
    painter.drawRoundedRect(right, 3.5, 3.5)
    painter.drawLine(int(mid), int(top + 1), int(mid), int(bottom - 1))

    painter.end()
    return pm


def _fallback_user_pixmap(color_hex: str = "#64748B", size: int = 18) -> QPixmap:
    """qtawesome 미설치/실패 시에도 항상 보이는 유저 아이콘."""
    from PyQt5.QtGui import QPainter, QPen, QBrush

    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)

    c = QColor(color_hex)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing, True)

    pen = QPen(c)
    pen.setWidthF(1.6)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(QBrush(Qt.NoBrush))

    # head
    r = size * 0.22
    cx = size * 0.5
    cy = size * 0.38
    painter.drawEllipse(int(cx - r), int(cy - r), int(2 * r), int(2 * r))
    # shoulders
    left = size * 0.22
    right = size * 0.78
    top = size * 0.62
    bottom = size * 0.92
    painter.drawRoundedRect(int(left), int(top), int(right - left), int(bottom - top), 3, 3)

    painter.end()
    return pm


def _brand_font(size_pt: int, weight: int) -> QFont:
    f = QFont("Inter")
    if not f.exactMatch():
        f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    f.setWeight(int(weight))
    # 세련된 로고 자간(타이트)
    try:
        f.setLetterSpacing(QFont.AbsoluteSpacing, -0.5)
    except Exception:
        pass
    return f


class Header(QFrame):
    """상단 헤더 위젯 (MainWindow와 호환 API 유지)"""

    tab_changed = pyqtSignal(str)  # 탭 변경 (수업준비, 수업, 관리)
    logo_clicked = pyqtSignal()  # 로고 클릭 (메인 페이지로 이동)
    profile_edit_clicked = pyqtSignal()  # 정보수정 클릭

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._tab_buttons: Dict[str, QPushButton] = {}
        self._init_ui()

    def _init_ui(self) -> None:
        self.setObjectName("Header")
        self.setFixedHeight(70)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        # ✅ 충돌 방지: 헤더 내부 스타일 초기화 + 새 표준만 적용
        self.setStyleSheet(
            """
            QFrame#Header {
                background-color: #FFFFFF;
                border-bottom: 1px solid #F1F5F9;
            }

            QFrame#Header * {
                border: none;
                background: transparent;
                outline: none;
                font-family: 'Pretendard','Inter','Malgun Gothic','맑은 고딕';
            }

            QLabel#LogoText {
                color: #2563EB;
            }

            QPushButton#MenuBtn {
                color: #64748B;
                font-size: 12pt;
                font-weight: 600;
                /* ✅ 심플 밑줄 스타일: 배경/박스 효과 제거 + 밑줄만 */
                padding: 10px 15px;
                padding-bottom: 12px; /* 텍스트-밑줄 간격 */
                border-bottom: 3px solid transparent;
                border-radius: 0px;
            }
            QPushButton#MenuBtn:hover {
                color: #2563EB;
            }
            QPushButton#MenuBtn[active="true"] {
                color: #2563EB;
                font-weight: 800;
                border-bottom: 3px solid #2563EB;
                border-radius: 0px;
            }

            QPushButton#IconBtn {
                border: none;
                border-radius: 8px;
                padding: 6px;
            }
            QPushButton#IconBtn:hover {
                background-color: #F8FAFC;
            }

            QLabel#UserInfo {
                color: #334155;
                font-size: 11pt;
                font-weight: 600;
            }
            QPushButton#LogoutBtn {
                color: #EF4444;
                font-size: 10pt;
                font-weight: 800;
                padding: 8px 10px;
                border-radius: 10px;
            }
            QPushButton#LogoutBtn:hover {
                color: #DC2626;
                background-color: #FEE2E2;
            }
            QPushButton#ProfileEditBtn {
                color: #64748B;
                font-size: 10pt;
                font-weight: 600;
                padding: 8px 12px;
                border-radius: 8px;
            }
            QPushButton#ProfileEditBtn:hover {
                color: #2563EB;
                background-color: #EFF6FF;
            }
            """
        )

        root = QHBoxLayout(self)
        root.setContentsMargins(32, 0, 30, 0)
        root.setSpacing(0)

        # 1) 로고 영역 (아이콘 + 텍스트, 10% 확대)
        logo_container = QFrame()
        logo_container.setObjectName("LogoContainer")
        logo_container.setCursor(Qt.PointingHandCursor)
        logo_container.setFocusPolicy(Qt.NoFocus)

        logo_layout = QHBoxLayout(logo_container)
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo_layout.setSpacing(8)

        logo_icon = QLabel()
        logo_icon.setFixedSize(26, 26)
        pm = _qta_pixmap("fa5s.graduation-cap", "#2563EB", 26)
        if pm is None:
            pm = _qta_pixmap("fa5s.book-reader", "#2563EB", 26)
        if pm is not None:
            logo_icon.setPixmap(pm)
        else:
            logo_icon.setPixmap(_fallback_logo_pixmap("#2563EB", 26))
            logo_icon.setText("")

        logo_text = QLabel("HanQ")
        logo_text.setObjectName("LogoText")
        logo_text.setFont(_brand_font(20, QFont.ExtraBold))

        logo_layout.addWidget(logo_icon, alignment=Qt.AlignVCenter)
        logo_layout.addWidget(logo_text, alignment=Qt.AlignVCenter)

        # 로고 클릭 (아이콘/텍스트/컨테이너 모두)
        logo_container.mousePressEvent = lambda event: self.logo_clicked.emit()
        logo_text.mousePressEvent = lambda event: self.logo_clicked.emit()
        logo_icon.mousePressEvent = lambda event: self.logo_clicked.emit()

        root.addWidget(logo_container, alignment=Qt.AlignVCenter)

        # 로고 ↔ 메뉴 간격(약 60px)
        root.addSpacing(60)

        # 2) 메뉴 탭
        self.tab_group = QButtonGroup(self)
        self.tab_group.setExclusive(True)

        for name in ["수업준비", "수업", "관리"]:
            btn = QPushButton(name)
            btn.setObjectName("MenuBtn")
            btn.setCheckable(True)
            btn.setFocusPolicy(Qt.NoFocus)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setProperty("active", False)

            # 버튼 width를 텍스트에 맞게 고정(underline이 과하게 길어지는 것 방지)
            btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

            btn.clicked.connect(lambda checked, n=name: self._on_tab_clicked(n))
            btn.toggled.connect(lambda checked, n=name: self._sync_tab_style(n, checked))

            self.tab_group.addButton(btn)
            self._tab_buttons[name] = btn
            root.addWidget(btn, alignment=Qt.AlignVCenter)
            root.addSpacing(48)

        # 회원관리 탭 (관리자 로그인 시에만 표시)
        self._members_btn = QPushButton("회원관리")
        self._members_btn.setObjectName("MenuBtn")
        self._members_btn.setCheckable(True)
        self._members_btn.setFocusPolicy(Qt.NoFocus)
        self._members_btn.setCursor(Qt.PointingHandCursor)
        self._members_btn.setProperty("active", False)
        self._members_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self._members_btn.setVisible(False)
        self._members_btn.clicked.connect(lambda: self._on_tab_clicked("회원관리"))
        self._members_btn.toggled.connect(lambda checked: self._sync_tab_style("회원관리", checked))
        self.tab_group.addButton(self._members_btn)
        self._tab_buttons["회원관리"] = self._members_btn
        root.addWidget(self._members_btn, alignment=Qt.AlignVCenter)
        root.addSpacing(48)

        # ✅ 메인 페이지 기본 상태: 어떤 탭도 선택하지 않음(하이라이트 없음)

        root.addStretch(1)

        # 3) 우측 섹션 (유저, 로그아웃)

        # 유저 정보 유닛(아이콘 + 문구)로 묶기
        user_unit = QWidget()
        user_unit.setFocusPolicy(Qt.NoFocus)
        user_layout = QHBoxLayout(user_unit)
        user_layout.setContentsMargins(0, 0, 0, 0)
        user_layout.setSpacing(8)
        user_layout.setAlignment(Qt.AlignVCenter)

        user_icon = QLabel()
        user_icon.setFocusPolicy(Qt.NoFocus)
        user_icon.setFixedSize(18, 18)
        pm = _qta_pixmap("fa5s.user-circle", "#64748B", 18)
        if pm is not None:
            user_icon.setPixmap(pm)
        else:
            user_icon.setPixmap(_fallback_user_pixmap("#64748B", 18))
        user_layout.addWidget(user_icon, alignment=Qt.AlignVCenter)

        self.user_info = QLabel("로그인해 주세요")
        self.user_info.setObjectName("UserInfo")
        self.user_info.setFont(_brand_font(11, QFont.DemiBold))
        user_layout.addWidget(self.user_info, alignment=Qt.AlignVCenter)

        self.profile_edit_btn = QPushButton("정보수정")
        self.profile_edit_btn.setObjectName("ProfileEditBtn")
        self.profile_edit_btn.setFocusPolicy(Qt.NoFocus)
        self.profile_edit_btn.setCursor(Qt.PointingHandCursor)
        self.profile_edit_btn.clicked.connect(self.profile_edit_clicked.emit)

        self.logout_btn = QPushButton("로그아웃")
        self.logout_btn.setObjectName("LogoutBtn")
        self.logout_btn.setFocusPolicy(Qt.NoFocus)
        self.logout_btn.setCursor(Qt.PointingHandCursor)

        root.addWidget(user_unit, alignment=Qt.AlignVCenter)
        root.addSpacing(16)
        root.addWidget(self.profile_edit_btn, alignment=Qt.AlignVCenter)
        root.addSpacing(16)
        root.addWidget(self.logout_btn, alignment=Qt.AlignVCenter)

    def set_user_display(self, name: str) -> None:
        """로그인한 사용자 이름 표시 (예: '홍길동님 환영합니다')"""
        self.user_info.setText(f"{name}님 환영합니다" if name else "로그인해 주세요")

    def set_admin_mode(self, visible: bool) -> None:
        """관리자일 때 회원관리 탭 표시/숨김"""
        if hasattr(self, "_members_btn") and self._members_btn is not None:
            self._members_btn.setVisible(visible)

    def _on_tab_clicked(self, tab_name: str) -> None:
        self.tab_changed.emit(tab_name)

    def clear_tab_selection(self) -> None:
        """어떤 탭도 선택되지 않은(하이라이트 없는) 상태로 초기화."""
        try:
            self.tab_group.setExclusive(False)
        except Exception:
            pass
        for btn in self._tab_buttons.values():
            try:
                btn.setChecked(False)
                btn.setProperty("active", False)
                btn.style().unpolish(btn)
                btn.style().polish(btn)
                btn.update()
            except Exception:
                continue
        try:
            self.tab_group.setExclusive(True)
        except Exception:
            pass

    def _sync_tab_style(self, tab_name: str, checked: bool) -> None:
        btn = self._tab_buttons.get(tab_name)
        if not btn:
            return
        btn.setProperty("active", bool(checked))
        btn.style().unpolish(btn)
        btn.style().polish(btn)
        btn.update()


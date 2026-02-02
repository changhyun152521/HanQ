"""
메인 페이지 화면 (Landing Page)

- 기존 디자인 전면 폐기 (Rewrite from scratch)
- 중앙 Hero + CTA + 3단 Feature 카드
- Opacity + Geometry 애니메이션
- 메인 페이지 진입 시 헤더 탭 하이라이트 해제
"""

from __future__ import annotations

from typing import Optional, List, Tuple

from PyQt5.QtCore import Qt, pyqtSignal, QTimer, QEasingCurve, QPropertyAnimation
from PyQt5.QtGui import QFont, QColor, QPixmap, QPainter, QPen
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QFrame,
    QSizePolicy,
    QGraphicsOpacityEffect,
)

try:
    import qtawesome as qta  # type: ignore

    _QTA_AVAILABLE = True
except Exception:
    qta = None
    _QTA_AVAILABLE = False


def _font(size_pt: int, weight: int) -> QFont:
    f = QFont("Inter")
    if not f.exactMatch():
        f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    f.setWeight(int(weight))
    return f


def _qta_pixmap(icon_name: str, color: str, size: int) -> Optional[QPixmap]:
    if not _QTA_AVAILABLE or qta is None:
        return None
    try:
        return qta.icon(icon_name, color=color).pixmap(size, size)
    except Exception:
        return None


def _fallback_file_pixmap(color_hex: str = "#2563EB", size: int = 28) -> QPixmap:
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    c = QColor(color_hex)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(c)
    pen.setWidthF(2.0)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    # sheet
    pad = 6
    painter.drawRoundedRect(pad, pad, size - pad * 2, size - pad * 2, 6, 6)
    # lines
    y1 = int(size * 0.45)
    y2 = int(size * 0.60)
    x1 = pad + 5
    x2 = size - pad - 5
    painter.drawLine(x1, y1, x2, y1)
    painter.drawLine(x1, y2, int(size * 0.78), y2)
    painter.end()
    return pm


def _fallback_check_pixmap(color_hex: str = "#2563EB", size: int = 28) -> QPixmap:
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    c = QColor(color_hex)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(c)
    pen.setWidthF(2.6)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    # check
    painter.drawLine(int(size * 0.20), int(size * 0.55), int(size * 0.42), int(size * 0.74))
    painter.drawLine(int(size * 0.42), int(size * 0.74), int(size * 0.82), int(size * 0.28))
    painter.end()
    return pm


def _fallback_user_pixmap(color_hex: str = "#2563EB", size: int = 28) -> QPixmap:
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    c = QColor(color_hex)
    painter = QPainter(pm)
    painter.setRenderHint(QPainter.Antialiasing, True)
    pen = QPen(c)
    pen.setWidthF(2.0)
    pen.setCapStyle(Qt.RoundCap)
    pen.setJoinStyle(Qt.RoundJoin)
    painter.setPen(pen)
    painter.setBrush(Qt.NoBrush)
    # head
    r = int(size * 0.22)
    cx = int(size * 0.5)
    cy = int(size * 0.38)
    painter.drawEllipse(cx - r, cy - r, 2 * r, 2 * r)
    # shoulders
    left = int(size * 0.20)
    top = int(size * 0.58)
    w = int(size * 0.60)
    h = int(size * 0.30)
    painter.drawRoundedRect(left, top, w, h, 6, 6)
    painter.end()
    return pm


def _main_stylesheet() -> str:
    return """
        QWidget#MainLandingRoot {
            background-color: #F8F9FA;
        }
        QWidget#MainLandingRoot * {
            outline: none;
            font-family: 'Pretendard','Inter','Malgun Gothic','맑은 고딕';
        }
        QWidget#MainLandingRoot QLabel {
            background: transparent;
            border: none;
        }

        QLabel#MainTitle {
            color: #1A1A1A;
            font-weight: 600;
        }
        QLabel#SubTitle {
            color: #555555;
        }

        QPushButton#StartButton {
            background-color: #3498DB;
            color: #FFFFFF;
            border: none;
            border-radius: 10px;
            padding: 12px 52px;
            font-weight: 800;
        }
        QPushButton#StartButton:hover {
            background-color: #2980B9;
        }
        QPushButton#StartButton:pressed {
            background-color: #2471A3;
        }

        QFrame#FeatureCard {
            background-color: #FFFFFF;
            border: 1px solid #E9ECEF;
            border-radius: 16px;
        }
        QLabel#CardTitle {
            color: #1A1A1A;
            font-weight: 600;
            border: none;
        }
        QLabel#CardDesc {
            color: #555555;
            border: none;
        }
    """


# 메인 컬러의 가장 진한 톤 (아이콘·그림자 가시성)
_CARD_ICON_COLOR = "#1D4ED8"


class FeatureCard(QFrame):
    def __init__(
        self,
        icon_key: str,
        title: str,
        desc: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setObjectName("FeatureCard")
        self.setFixedSize(300, 190)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setAutoFillBackground(True)

        self._radius = 20

        # QGraphicsDropShadowEffect는 Opacity/Geometry 애니메이션과 충돌하여
        # QPainter 오류 및 카드 소멸을 유발하므로 제거. 그림자는 QSS border로 대체.
        root = QVBoxLayout(self)
        root.setContentsMargins(31, 27, 31, 27)
        root.setSpacing(10)

        icon = QLabel()
        icon.setFixedSize(34, 34)
        icon.setStyleSheet("border: none;")
        pm = None
        if icon_key == "file":
            pm = _qta_pixmap("fa5s.file-alt", _CARD_ICON_COLOR, 32)
            if pm is None:
                pm = _fallback_file_pixmap(_CARD_ICON_COLOR, 32)
        elif icon_key == "check":
            pm = _qta_pixmap("fa5s.check-circle", _CARD_ICON_COLOR, 32)
            if pm is None:
                pm = _fallback_check_pixmap(_CARD_ICON_COLOR, 32)
        elif icon_key == "user":
            pm = _qta_pixmap("fa5s.user-graduate", _CARD_ICON_COLOR, 32)
            if pm is None:
                pm = _fallback_user_pixmap(_CARD_ICON_COLOR, 32)
        if pm is not None:
            icon.setPixmap(pm)
        root.addWidget(icon, alignment=Qt.AlignLeft)

        t = QLabel(title)
        t.setObjectName("CardTitle")
        t.setFont(_font(12, QFont.DemiBold))
        root.addWidget(t)

        d = QLabel(desc)
        d.setObjectName("CardDesc")
        d.setFont(_font(10, QFont.Medium))
        d.setWordWrap(True)
        root.addWidget(d)
        root.addStretch(1)

    def paintEvent(self, event):  # type: ignore[override]
        super().paintEvent(event)


class CtaButton(QPushButton):
    """그라데이션 + 소프트 쉐도우를 graphicsEffect 없이 직접 렌더링."""

    def __init__(self, text: str, parent: Optional[QWidget] = None):
        super().__init__(text, parent)
        self.setAttribute(Qt.WA_StyledBackground, True)

    def paintEvent(self, event):  # type: ignore[override]
        # 레거시 커스텀 페인팅 제거(플랫 버튼은 QSS로 처리)
        super().paintEvent(event)


class MainPageScreen(QWidget):
    """메인 랜딩 페이지"""

    start_requested = pyqtSignal()

    def __init__(self, db_connection=None, parent=None):
        super().__init__(parent)
        self.db_connection = db_connection  # 구조 유지(미사용)
        self._anim_refs: List[QPropertyAnimation] = []
        self._animated_once = False
        self._build_ui()

    @staticmethod
    def _wrap_for_anim(child: QWidget) -> QWidget:
        """DropShadow 등 graphicsEffect가 있는 위젯도 Opacity 애니메이션 가능하게 래핑."""
        w = QWidget()
        w.setContentsMargins(0, 0, 0, 0)
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(child)
        # child가 고정 크기면 wrapper도 동일하게
        try:
            w.setFixedSize(child.size())
        except Exception:
            pass
        return w

    def _build_ui(self):
        self.setObjectName("MainLandingRoot")
        self.setStyleSheet(_main_stylesheet())

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # ✅ 반응형 중앙 정렬(Vertical Centering): 위/아래 stretch
        outer.addStretch(1)

        container = QWidget()
        container.setMaximumWidth(1120)
        container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        container_layout = QVBoxLayout(container)
        # ✅ 그림자 잘림 방지: 중앙 컨테이너 자체에 여백을 줘서 안전 공간 확보
        container_layout.setContentsMargins(20, 20, 20, 20)
        container_layout.setSpacing(0)
        container_layout.setAlignment(Qt.AlignHCenter)

        # Hero (타이틀 + 설명)
        hero = QWidget()
        hero_layout = QVBoxLayout(hero)
        hero_layout.setContentsMargins(0, 0, 0, 0)
        hero_layout.setSpacing(16)
        hero_layout.setAlignment(Qt.AlignCenter)

        self.title = QLabel("HanQ")
        self.title.setObjectName("MainTitle")
        self.title.setAlignment(Qt.AlignCenter)
        f = _font(32, QFont.DemiBold)
        try:
            f.setLetterSpacing(QFont.AbsoluteSpacing, -0.6)
        except Exception:
            pass
        self.title.setFont(f)
        hero_layout.addWidget(self.title)

        # ✅ 줄 길이/줄간격(160%)을 강제하기 위해 RichText 사용
        self.desc = QLabel(
            "<div style='text-align:center; line-height:160%;'>"
            "한글문서 기반의 학생별 맞춤 학습지부터<br/>"
            "오답노트 및 개별 보고서 생성까지</div>"
        )
        self.desc.setObjectName("SubTitle")
        self.desc.setAlignment(Qt.AlignCenter)
        self.desc.setTextFormat(Qt.RichText)
        self.desc.setFont(_font(15, QFont.Medium))
        hero_layout.addWidget(self.desc)

        container_layout.addWidget(hero)
        container_layout.addSpacing(40)

        # CTA (강조 + 여백)
        self.cta_btn = QPushButton("시작하기")
        self.cta_btn.setObjectName("StartButton")
        self.cta_btn.setCursor(Qt.PointingHandCursor)
        self.cta_btn.setFocusPolicy(Qt.NoFocus)
        self.cta_btn.setFont(_font(15, QFont.Bold))
        self.cta_btn.setFixedHeight(46)
        self.cta_btn.setMinimumWidth(220)
        self.cta_btn.clicked.connect(self.start_requested.emit)

        self.cta_wrap = self._wrap_for_anim(self.cta_btn)
        # ✅ 버튼 shadow 잘림 방지: 버튼을 담는 레이아웃에 여백(20px)
        cta_section = QWidget()
        cta_lay = QHBoxLayout(cta_section)
        cta_lay.setContentsMargins(20, 20, 20, 20)
        cta_lay.setSpacing(0)
        cta_lay.addStretch(1)
        cta_lay.addWidget(self.cta_wrap)
        cta_lay.addStretch(1)
        container_layout.addWidget(cta_section)
        # ✅ 카드 섹션과 최소 60px 이상 간격
        container_layout.addSpacing(80)

        # Feature cards (3단)
        cards_wrap = QWidget()
        cards_layout = QHBoxLayout(cards_wrap)
        # ✅ 카드 shadow 잘림 방지: 카드들을 담는 레이아웃에 여백(20px)
        cards_layout.setContentsMargins(20, 20, 20, 20)
        cards_layout.setSpacing(24)
        cards_layout.setAlignment(Qt.AlignCenter)

        self.card_1 = FeatureCard("file", "hwp 문항 관리", "한글 파일을 기반으로 한 정밀한 DB 구축")
        self.card_2 = FeatureCard("check", "맞춤형 학습지", "클릭 몇 번으로 완성되는 학생별 맞춤 프린트")
        self.card_3 = FeatureCard("user", "성적 & 오답 관리", "학생별 오답 데이터를 분석하여 자동으로 생성되는 오답노트")

        self.card_1_wrap = self._wrap_for_anim(self.card_1)
        self.card_2_wrap = self._wrap_for_anim(self.card_2)
        self.card_3_wrap = self._wrap_for_anim(self.card_3)
        cards_layout.addWidget(self.card_1_wrap)
        cards_layout.addWidget(self.card_2_wrap)
        cards_layout.addWidget(self.card_3_wrap)

        container_layout.addWidget(cards_wrap, alignment=Qt.AlignHCenter)

        outer.addWidget(container, alignment=Qt.AlignHCenter)
        outer.addStretch(1)
        outer.addWidget(self._create_footer())

    def _create_footer(self) -> QWidget:
        """메인 페이지 최하단 고정형 푸터(저작권·개발자 정보). 시인성 강화."""
        footer_widget = QWidget()
        footer_layout = QVBoxLayout(footer_widget)
        footer_layout.setContentsMargins(0, 0, 0, 25)
        footer_layout.setSpacing(2)

        copyright_lbl = QLabel("© 2026 HanQ. All rights reserved.")
        copyright_lbl.setAlignment(Qt.AlignCenter)
        copyright_lbl.setStyleSheet(
            "color: #777777; font-size: 10pt; font-weight: 400; border: none;"
        )
        copyright_lbl.setFont(_font(10, QFont.Normal))

        dev_lbl = QLabel(
            'Developed by <span style="font-weight: bold; color: #333333;">이창현수학</span>'
        )
        dev_lbl.setAlignment(Qt.AlignCenter)
        dev_lbl.setTextFormat(Qt.RichText)
        dev_lbl.setStyleSheet(
            "color: #666666; font-size: 10pt; border: none;"
        )
        dev_lbl.setFont(_font(10, QFont.Normal))

        footer_layout.addWidget(copyright_lbl)
        footer_layout.addWidget(dev_lbl)

        return footer_widget

    def showEvent(self, event):  # type: ignore[override]
        super().showEvent(event)

        # 1) 메인 페이지 진입 시: 헤더 탭 하이라이트 완전 해제
        try:
            w = self.window()
            header = getattr(w, "header", None)
            if header is not None:
                # Header에 전용 API가 있으면 그걸 사용(Exclusive 그룹에서도 확실히)
                if hasattr(header, "clear_tab_selection"):
                    header.clear_tab_selection()
                elif getattr(header, "tab_group", None) is not None:
                    for btn in header.tab_group.buttons():
                        btn.setChecked(False)
        except Exception:
            pass

        # 2) 애니메이션 (Opacity + Geometry)
        if self._animated_once:
            return
        self._animated_once = True
        QTimer.singleShot(30, self._run_animations)

    def hideEvent(self, event):  # type: ignore[override]
        # 화면 전환 후 다시 들어왔을 때도 "열릴 때" 효과가 필요하면 재생
        self._animated_once = False
        super().hideEvent(event)

    def _run_animations(self):
        self._anim_refs.clear()

        targets: List[Tuple[QWidget, int]] = [
            (self.title, 0),
            (self.desc, 60),
            (self.cta_wrap, 120),
            (self.card_1_wrap, 180),
            (self.card_2_wrap, 240),
            (self.card_3_wrap, 300),
        ]

        for w, delay in targets:
            self._animate_in(w, delay_ms=delay)

    def _animate_in(self, w: QWidget, delay_ms: int):
        try:
            end_geo = w.geometry()
            start_geo = w.geometry()
            start_geo.moveTop(start_geo.top() + 18)

            eff = QGraphicsOpacityEffect(w)
            w.setGraphicsEffect(eff)
            eff.setOpacity(0.0)

            geo_anim = QPropertyAnimation(w, b"geometry", self)
            geo_anim.setDuration(520)
            geo_anim.setStartValue(start_geo)
            geo_anim.setEndValue(end_geo)
            geo_anim.setEasingCurve(QEasingCurve.OutCubic)

            self._anim_refs.append(geo_anim)

            op_anim = QPropertyAnimation(eff, b"opacity", self)
            op_anim.setDuration(520)
            op_anim.setStartValue(0.0)
            op_anim.setEndValue(1.0)
            op_anim.setEasingCurve(QEasingCurve.OutCubic)
            self._anim_refs.append(op_anim)

            def _start():
                geo_anim.start()
                op_anim.start()

            QTimer.singleShot(max(0, int(delay_ms)), _start)
        except Exception:
            return

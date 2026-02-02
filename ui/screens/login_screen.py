"""
로그인 화면

- 정중앙 배치, 단정한 박스 스타일 (로고 / 아이디 / 비밀번호 / 로그인 버튼)
- 포커스 효과 없음, 푸터(저작권·개발자) 하단 고정
"""
from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
    QFrame,
    QGraphicsDropShadowEffect,
)


def _font(size_pt: int, weight: int = QFont.Normal) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    f.setWeight(int(weight))
    return f


class LoginScreen(QWidget):
    """로그인 화면 (박스 스타일 입력창, 포커스 없음, 하단 푸터)"""

    login_succeeded = pyqtSignal(str, str)  # user_id, name

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self._build_ui()

    def _build_ui(self) -> None:
        self.setObjectName("LoginScreen")
        self.setStyleSheet("QWidget#LoginScreen { background-color: #FFFFFF; }")

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 20)
        main_layout.addStretch(1)

        # 로그인 카드 (가로 380px, 수평 중앙)
        self.login_card = QFrame()
        self.login_card.setObjectName("LoginCard")
        self.login_card.setFixedWidth(380)
        self.login_card.setStyleSheet("""
            QFrame#LoginCard {
                background-color: white;
                border: none;
            }
            QFrame#LoginCard QLabel#LoginTitle {
                font-size: 36pt;
                font-weight: 900;
                color: #000000;
                background: transparent;
                border: none;
            }
            QFrame#LoginCard QLineEdit {
                background-color: white;
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                padding: 12px;
                font-size: 11pt;
                color: #1A1A1A;
            }
            QFrame#LoginCard QLineEdit:focus {
                border: 1px solid #E0E0E0;
            }
            QFrame#LoginCard QLineEdit::placeholder {
                color: #555555;
            }
            QFrame#LoginCard QPushButton#LoginBtn {
                background-color: #1976D2;
                color: white;
                font-weight: bold;
                font-size: 13pt;
                border: none;
                border-radius: 8px;
                padding: 12px;
            }
            QFrame#LoginCard QPushButton#LoginBtn:hover {
                background-color: #1565C0;
            }
            QFrame#LoginCard QPushButton#LoginBtn:pressed {
                background-color: #0D47A1;
            }
            QFrame#LoginCard QPushButton#LoginBtn:disabled {
                background-color: #B0BEC5;
            }
        """)

        shadow = QGraphicsDropShadowEffect(self.login_card)
        shadow.setBlurRadius(20)
        shadow.setXOffset(0)
        shadow.setYOffset(2)
        shadow.setColor(QColor(0, 0, 0, 30))
        self.login_card.setGraphicsEffect(shadow)

        card_layout = QVBoxLayout(self.login_card)
        card_layout.setContentsMargins(40, 36, 40, 36)
        card_layout.setSpacing(24)
        card_layout.setAlignment(Qt.AlignCenter)

        # 제목 HanQ (배경·테두리 없이 글자만)
        title = QLabel("HanQ")
        title.setObjectName("LoginTitle")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(_font(36, QFont.Black))
        title.setStyleSheet("background: transparent; border: none; color: #000000;")
        card_layout.addWidget(title)

        card_layout.addSpacing(16)

        # 입력창: 둥근 박스, 흰 배경, 연한 회색 테두리, 포커스 시에도 동일
        self.inp_id = QLineEdit()
        self.inp_id.setPlaceholderText("아이디를 입력하세요")
        self.inp_id.setMinimumHeight(44)
        self.inp_id.setClearButtonEnabled(True)
        card_layout.addWidget(self.inp_id)

        self.inp_password = QLineEdit()
        self.inp_password.setPlaceholderText("비밀번호를 입력하세요")
        self.inp_password.setEchoMode(QLineEdit.Password)
        self.inp_password.setMinimumHeight(44)
        self.inp_password.setClearButtonEnabled(False)
        card_layout.addWidget(self.inp_password)

        card_layout.addSpacing(12)

        self.btn_login = QPushButton("로그인")
        self.btn_login.setObjectName("LoginBtn")
        self.btn_login.setCursor(Qt.PointingHandCursor)
        self.btn_login.setFocusPolicy(Qt.NoFocus)
        self.btn_login.setMinimumHeight(48)
        self.btn_login.clicked.connect(self._on_login)
        self.inp_password.returnPressed.connect(self._on_login)
        card_layout.addWidget(self.btn_login)

        # 카드 수평 중앙 배치
        h_layout = QHBoxLayout()
        h_layout.addStretch(1)
        h_layout.addWidget(self.login_card)
        h_layout.addStretch(1)
        main_layout.addLayout(h_layout)

        main_layout.addStretch(1)

        # 푸터 (메인 페이지 스타일, 바닥 고정)
        footer_layout = QVBoxLayout()
        footer_layout.setSpacing(4)

        copy_lbl = QLabel("© 2026 HanQ. All rights reserved.")
        copy_lbl.setStyleSheet("color: #777777; font-size: 10pt; background: transparent; border: none;")
        copy_lbl.setAlignment(Qt.AlignCenter)

        dev_lbl = QLabel("Developed by 이창현수학")
        dev_lbl.setStyleSheet("color: #333333; font-size: 10pt; font-weight: bold; background: transparent; border: none;")
        dev_lbl.setAlignment(Qt.AlignCenter)
        dev_lbl.setFont(_font(10, QFont.DemiBold))

        footer_layout.addWidget(copy_lbl)
        footer_layout.addWidget(dev_lbl)
        main_layout.addLayout(footer_layout)

    def _on_login(self) -> None:
        user_id = (self.inp_id.text() or "").strip()
        password = self.inp_password.text() or ""

        if not user_id:
            QMessageBox.information(self, "입력 필요", "아이디를 입력해 주세요.")
            self.inp_id.setFocus()
            return

        self.btn_login.setEnabled(False)
        try:
            from services.login_api import login as api_login
            result = api_login(user_id, password)
        except Exception as e:
            QMessageBox.warning(self, "오류", f"로그인 요청 중 오류가 발생했습니다.\n\n{e}")
            self.btn_login.setEnabled(True)
            return

        self.btn_login.setEnabled(True)
        if result.get("success"):
            name = result.get("name") or user_id
            self.login_succeeded.emit(user_id, name)
        else:
            QMessageBox.warning(self, "로그인 실패", result.get("message") or "아이디 또는 비밀번호를 확인해 주세요.")

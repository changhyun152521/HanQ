"""
정보수정 다이얼로그

- 이름, 아이디, 비밀번호 변경 (현재 비밀번호 확인 필요)
"""
from __future__ import annotations

from typing import Optional

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QMessageBox,
)


def _font(size_pt: int, weight: int = QFont.Normal) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    f.setWeight(int(weight))
    return f


class ProfileEditDialog(QDialog):
    """이름·아이디·비밀번호 수정 다이얼로그"""

    def __init__(
        self,
        current_user_id: str,
        current_name: str,
        parent: Optional[QDialog] = None,
    ):
        super().__init__(parent)
        self.current_user_id = (current_user_id or "").strip()
        self.current_name = (current_name or current_user_id or "").strip()
        self._build_ui()

    def _build_ui(self) -> None:
        self.setWindowTitle("정보 수정")
        self.setMinimumWidth(360)
        self.setStyleSheet("""
            QDialog {
                background-color: #FFFFFF;
            }
            QLabel {
                background: none;
                background-color: transparent;
                border: none;
                color: #0F172A;
                font-weight: 700;
            }
            QLineEdit {
                background-color: #FFFFFF;
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                padding: 10px 12px;
                font-size: 11pt;
                font-weight: 600;
                color: #1A1A1A;
            }
            QPushButton {
                border-radius: 8px;
                padding: 10px 16px;
                font-weight: 700;
            }
            QPushButton#OkBtn {
                background-color: #1976D2;
                color: white;
            }
            QPushButton#OkBtn:hover { background-color: #1565C0; }
            QPushButton#CancelBtn {
                background-color: #F1F5F9;
                color: #64748B;
            }
            QPushButton#CancelBtn:hover { background-color: #E2E8F0; }
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("회원 정보 수정")
        title.setFont(_font(15, QFont.ExtraBold))
        title.setStyleSheet("background: none; background-color: transparent; border: none;")
        layout.addWidget(title)

        layout.addSpacing(8)

        # 현재 비밀번호 (필수)
        lbl_password = QLabel("현재 비밀번호")
        lbl_password.setFont(_font(11, QFont.Bold))
        layout.addWidget(lbl_password)
        self.inp_password = QLineEdit()
        self.inp_password.setPlaceholderText("현재 비밀번호 입력")
        self.inp_password.setEchoMode(QLineEdit.Password)
        self.inp_password.setMinimumHeight(40)
        layout.addWidget(self.inp_password)

        lbl_name = QLabel("이름")
        lbl_name.setFont(_font(11, QFont.Bold))
        layout.addWidget(lbl_name)
        self.inp_name = QLineEdit()
        self.inp_name.setPlaceholderText("이름")
        self.inp_name.setText(self.current_name)
        self.inp_name.setMinimumHeight(40)
        layout.addWidget(self.inp_name)

        lbl_user_id = QLabel("아이디")
        lbl_user_id.setFont(_font(11, QFont.Bold))
        layout.addWidget(lbl_user_id)
        self.inp_user_id = QLineEdit()
        self.inp_user_id.setPlaceholderText("아이디")
        self.inp_user_id.setText(self.current_user_id)
        self.inp_user_id.setMinimumHeight(40)
        layout.addWidget(self.inp_user_id)

        lbl_new_pw = QLabel("새 비밀번호 (변경 시에만 입력)")
        lbl_new_pw.setFont(_font(11, QFont.Bold))
        layout.addWidget(lbl_new_pw)
        self.inp_new_password = QLineEdit()
        self.inp_new_password.setPlaceholderText("비워두면 기존 비밀번호 유지")
        self.inp_new_password.setEchoMode(QLineEdit.Password)
        self.inp_new_password.setMinimumHeight(40)
        layout.addWidget(self.inp_new_password)

        layout.addSpacing(12)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch(1)
        cancel_btn = QPushButton("취소")
        cancel_btn.setObjectName("CancelBtn")
        cancel_btn.setFocusPolicy(Qt.NoFocus)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        ok_btn = QPushButton("저장")
        ok_btn.setObjectName("OkBtn")
        ok_btn.setFocusPolicy(Qt.NoFocus)
        ok_btn.clicked.connect(self._on_save)
        btn_layout.addWidget(ok_btn)
        layout.addLayout(btn_layout)

    def _on_save(self) -> None:
        password = self.inp_password.text() or ""
        if not password:
            QMessageBox.warning(self, "입력 필요", "현재 비밀번호를 입력해 주세요.")
            self.inp_password.setFocus()
            return

        new_name = (self.inp_name.text() or "").strip()
        new_user_id = (self.inp_user_id.text() or "").strip()
        new_password = (self.inp_new_password.text() or "").strip() or None

        from services.login_api import update_user

        result = update_user(
            self.current_user_id,
            password,
            new_user_id=new_user_id if new_user_id else None,
            new_name=new_name if new_name else None,
            new_password=new_password,
        )

        if not result.get("success"):
            QMessageBox.warning(self, "수정 실패", result.get("message", "정보 수정에 실패했습니다."))
            return

        self._result_user_id = result.get("user_id") or self.current_user_id
        self._result_name = result.get("name") or new_name or self.current_name
        self.accept()

    def get_updated_profile(self) -> tuple[str, str]:
        """저장 성공 시 (user_id, name) 반환. accept() 후에만 유효."""
        return (
            getattr(self, "_result_user_id", self.current_user_id),
            getattr(self, "_result_name", self.current_name),
        )

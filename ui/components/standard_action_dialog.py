"""
표준 액션 모달(화이트 톤)

- 전역 테마/다크 테마 영향을 받더라도, 이 다이얼로그 내부는 항상 밝은 톤으로 보이도록
  QSS를 강하게 오버라이드합니다.
- "앞으로 만들어지는 모든 모달"의 기본 골격으로 재사용할 수 있게 만들었습니다.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QDialog, QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget


@dataclass(frozen=True)
class DialogAction:
    key: str
    label: str
    is_primary: bool = False


class StandardActionDialog(QDialog):
    """
    깔끔한 화이트 톤 표준 모달.

    exec_() 후 selected_key로 어떤 버튼이 눌렸는지 확인합니다.
    """

    def __init__(
        self,
        *,
        title: str,
        message: str,
        actions: List[DialogAction],
        parent: Optional[QWidget] = None,
        min_width: int = 340,
    ):
        super().__init__(parent)
        self.setWindowTitle(title or "")
        self.setModal(True)
        self.selected_key: Optional[str] = None

        # 윈도우 기본 프레임/타이틀바를 유지한 채, 내용 영역만 화이트 톤으로 강제
        self.setAttribute(Qt.WA_StyledBackground, True)

        self._actions = list(actions or [])
        self._init_ui(message=message, min_width=int(min_width))

    def _init_ui(self, *, message: str, min_width: int) -> None:
        self.setMinimumWidth(max(280, int(min_width)))

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 16)
        root.setSpacing(14)

        msg = QLabel(message or "")
        msg.setWordWrap(True)
        msg.setObjectName("Message")
        root.addWidget(msg)

        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(0, 6, 0, 0)
        btn_row.setSpacing(10)
        btn_row.addStretch(1)

        # 우측 정렬, primary는 강조
        for a in self._actions:
            b = QPushButton(a.label)
            b.setCursor(Qt.PointingHandCursor)
            b.setFocusPolicy(Qt.NoFocus)
            b.setMinimumWidth(86)
            b.setMinimumHeight(34)
            b.setObjectName("PrimaryButton" if a.is_primary else "SecondaryButton")
            b.clicked.connect(lambda _checked=False, k=a.key: self._choose(k))
            btn_row.addWidget(b)

        root.addLayout(btn_row)

        # ✅ 표준 모달 QSS (전역 테마/다크 영향 차단용으로 강하게 지정)
        self.setStyleSheet(
            """
            QDialog {
                background-color: #FFFFFF;
            }
            QLabel#Message {
                color: #1E293B;
                font-size: 12px;
                font-weight: 600;
                background: transparent;
            }
            QPushButton#SecondaryButton {
                background-color: #F8FAFC;
                border: 1px solid #CBD5E1;
                border-radius: 8px;
                padding: 8px 14px;
                color: #1E293B;
                font-weight: 700;
            }
            QPushButton#SecondaryButton:hover {
                background-color: #F1F5F9;
                border-color: #94A3B8;
            }
            QPushButton#PrimaryButton {
                background-color: #2563EB;
                border: 1px solid #2563EB;
                border-radius: 8px;
                padding: 8px 14px;
                color: #FFFFFF;
                font-weight: 800;
            }
            QPushButton#PrimaryButton:hover {
                background-color: #1D4ED8;
                border-color: #1D4ED8;
            }
            """
        )

    def _choose(self, key: str) -> None:
        self.selected_key = key
        self.accept()


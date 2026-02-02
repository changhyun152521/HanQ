"""
표준 메시지/확인 모달 유틸

- QMessageBox 대신 StandardActionDialog(화이트 톤)를 사용해서
  앱 전반의 "모달" 디자인을 통일합니다.
"""

from __future__ import annotations

from typing import Optional

from PyQt5.QtWidgets import QWidget

from ui.components.standard_action_dialog import DialogAction, StandardActionDialog


def show_info(parent: Optional[QWidget], title: str, message: str) -> None:
    dlg = StandardActionDialog(
        parent=parent,
        title=title,
        message=message,
        actions=[DialogAction(key="ok", label="확인", is_primary=True)],
        min_width=380,
    )
    dlg.exec_()


def show_warning(parent: Optional[QWidget], title: str, message: str) -> None:
    dlg = StandardActionDialog(
        parent=parent,
        title=title,
        message=message,
        actions=[DialogAction(key="ok", label="확인", is_primary=True)],
        min_width=380,
    )
    dlg.exec_()


def confirm(parent: Optional[QWidget], title: str, message: str, *, ok_label: str = "확인", cancel_label: str = "취소") -> bool:
    dlg = StandardActionDialog(
        parent=parent,
        title=title,
        message=message,
        actions=[
            DialogAction(key="cancel", label=cancel_label, is_primary=False),
            DialogAction(key="ok", label=ok_label, is_primary=True),
        ],
        min_width=420,
    )
    dlg.exec_()
    return dlg.selected_key == "ok"


"""
채점 모달

요구사항:
- 출제된 학습지의 총 문항수만큼 번호별 O/X 선택
- 전체O / 전체X / 전체해제
- 모두 선택 후 "등록"하면 DB에 저장(몇개중 몇개 + 단원별 통계 + 번호별 결과)
"""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.models import Worksheet, Problem
from ui.components.standard_message import show_warning


def _font(size_pt: int, *, bold: bool = False, extra_bold: bool = False) -> QFont:
    f = QFont("Pretendard")
    if not f.exactMatch():
        f = QFont("맑은 고딕")
    f.setPointSize(int(size_pt))
    if extra_bold:
        f.setWeight(QFont.ExtraBold)
    elif bold:
        f.setBold(True)
    else:
        try:
            f.setWeight(QFont.Medium)
        except Exception:
            pass
    return f


def _problem_unit_key(p: Optional[Problem]) -> Tuple[str, str, str]:
    """
    returns (unit_key, major_unit, sub_unit)
    """
    if not p or not getattr(p, "tags", None):
        return ("미분류", "", "")
    tag0 = None
    for t in (p.tags or []):
        if getattr(t, "major_unit", None) or getattr(t, "unit", None):
            tag0 = t
            break
    if tag0 is None:
        return ("미분류", "", "")
    major = (getattr(tag0, "major_unit", None) or "").strip()
    sub = (getattr(tag0, "sub_unit", None) or "").strip()
    unit_raw = (getattr(tag0, "unit", None) or "").strip()
    if major and sub:
        return (f"{major} > {sub}", major, sub)
    if major:
        return (major, major, sub or "")
    if unit_raw:
        return (unit_raw, "", "")
    return ("미분류", "", "")


def _extract_problem_id(d: dict) -> str:
    """
    numbered/answer dict에서 problem_id를 최대한 복구합니다.
    (레거시 키 호환: problemId 등)
    """
    if not isinstance(d, dict):
        return ""
    for k in ("problem_id", "problemId", "problemID", "pid", "problem"):
        v = d.get(k)
        if v is None:
            continue
        s = str(v).strip()
        if s:
            return s
    return ""


class GradingDialog(QDialog):
    """
    결과는 다음 형태로 반환:
    - answers: [{no, problem_id, is_correct, unit_key, major_unit, sub_unit}, ...]
    - total_questions, correct_count
    - unit_stats: {unit_key: {total, correct}, ...}
    """

    def __init__(
        self,
        *,
        worksheet: Worksheet,
        numbered: List[dict],
        problems_by_id: Dict[str, Problem],
        existing_answers: Optional[Dict[int, bool]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self.worksheet = worksheet
        self.numbered = list(numbered or [])
        self.problems_by_id = dict(problems_by_id or {})

        self._answers: Dict[int, Optional[bool]] = {}  # no -> True(O)/False(X)/None
        if existing_answers:
            for k, v in existing_answers.items():
                self._answers[int(k)] = bool(v)

        # default: None
        for it in self.numbered:
            try:
                no = int(it.get("no"))
            except Exception:
                continue
            self._answers.setdefault(no, None)

        self.setWindowTitle("채점")
        self.setModal(True)
        self.setMinimumWidth(640)
        self.setMinimumHeight(680)
        self._build_ui()
        self._sync_summary()

    def result_payload(self) -> dict:
        total = len([k for k in self._answers.keys()])
        correct = sum(1 for k, v in self._answers.items() if v is True)

        answers: List[Dict] = []
        unit_stats: Dict[str, Dict] = {}
        for it in self.numbered:
            try:
                no = int(it.get("no"))
            except Exception:
                continue
            pid = _extract_problem_id(it)
            v = self._answers.get(no, None)
            is_correct = bool(v is True)
            p = self.problems_by_id.get(pid)
            unit_key, major, sub = _problem_unit_key(p)

            answers.append(
                {
                    "no": no,
                    "problem_id": pid,
                    "is_correct": bool(v is True),
                    "unit_key": unit_key,
                    "major_unit": major,
                    "sub_unit": sub,
                }
            )
            st = unit_stats.setdefault(unit_key, {"total": 0, "correct": 0})
            st["total"] = int(st.get("total", 0)) + 1
            st["correct"] = int(st.get("correct", 0)) + (1 if is_correct else 0)

        return {
            "total_questions": int(total),
            "correct_count": int(correct),
            "answers": answers,
            "unit_stats": unit_stats,
        }

    def _build_ui(self) -> None:
        self.setObjectName("GradingDialog")
        self.setStyleSheet(
            """
            QDialog#GradingDialog { background: #FFFFFF; }

            QLabel#Title { color:#0F172A; }
            QLabel#Sub { color:#64748B; }

            QFrame#Toolbar {
                background: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
            }
            QPushButton#ToolBtn {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 10px;
                padding: 8px 12px;
                font-weight: 900;
                color: #1E293B;
            }
            QPushButton#ToolBtn:hover {
                border-color:#2563EB;
                color:#2563EB;
                background:#F0F7FF;
            }

            QPushButton#NoBtn {
                background: #FFFFFF;
                border: 1px solid #CBD5E1;
                border-radius: 12px;
                padding: 10px 0px;
                font-weight: 900;
                color: #0F172A;
            }
            QPushButton#NoBtn[mark="O"] {
                background: #DCFCE7;
                border-color: #22C55E;
                color: #166534;
            }
            QPushButton#NoBtn[mark="X"] {
                background: #FEE2E2;
                border-color: #EF4444;
                color: #991B1B;
            }
            QPushButton#NoBtn:hover {
                border-color: #2563EB;
            }

            QPushButton#PrimaryBtn {
                background: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 900;
            }
            QPushButton#PrimaryBtn:hover { background:#1D4ED8; }
            QPushButton#GhostBtn {
                background: #F1F5F9;
                color: #1E293B;
                border: 1px solid #E2E8F0;
                border-radius: 12px;
                padding: 10px 16px;
                font-weight: 800;
            }
            QPushButton#GhostBtn:hover { background:#E2E8F0; }
            """
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(12)

        title = QLabel("채점")
        title.setObjectName("Title")
        title.setFont(_font(14, extra_bold=True))
        root.addWidget(title)

        ws_title = (self.worksheet.title or "").strip()
        sub = QLabel(ws_title)
        sub.setObjectName("Sub")
        sub.setFont(_font(10, bold=True))
        root.addWidget(sub)

        # toolbar
        tb = QFrame()
        tb.setObjectName("Toolbar")
        tb_lay = QHBoxLayout(tb)
        tb_lay.setContentsMargins(12, 10, 12, 10)
        tb_lay.setSpacing(10)

        self.lbl_summary = QLabel("")
        self.lbl_summary.setFont(_font(10, extra_bold=True))
        self.lbl_summary.setStyleSheet("color:#0F172A;")
        tb_lay.addWidget(self.lbl_summary, alignment=Qt.AlignVCenter)
        tb_lay.addStretch(1)

        btn_all_o = QPushButton("전체 O")
        btn_all_o.setObjectName("ToolBtn")
        btn_all_o.setCursor(Qt.PointingHandCursor)
        btn_all_o.clicked.connect(lambda: self._set_all(True))
        tb_lay.addWidget(btn_all_o)

        btn_all_x = QPushButton("전체 X")
        btn_all_x.setObjectName("ToolBtn")
        btn_all_x.setCursor(Qt.PointingHandCursor)
        btn_all_x.clicked.connect(lambda: self._set_all(False))
        tb_lay.addWidget(btn_all_x)

        btn_clear = QPushButton("전체 해제")
        btn_clear.setObjectName("ToolBtn")
        btn_clear.setCursor(Qt.PointingHandCursor)
        btn_clear.clicked.connect(self._clear_all)
        tb_lay.addWidget(btn_clear)

        root.addWidget(tb)

        # grid (scroll)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        root.addWidget(scroll, 1)

        cont = QWidget()
        scroll.setWidget(cont)

        grid = QGridLayout(cont)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        # 8 columns
        cols = 8
        nos = sorted(self._answers.keys())
        self._no_buttons: Dict[int, QPushButton] = {}
        for idx, no in enumerate(nos):
            r = idx // cols
            c = idx % cols
            btn = QPushButton(str(no))
            btn.setObjectName("NoBtn")
            btn.setFixedHeight(46)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setFont(_font(10, extra_bold=True))
            btn.clicked.connect(lambda _=False, n=no: self._toggle_no(n))
            self._no_buttons[no] = btn
            grid.addWidget(btn, r, c)

        # bottom
        bottom = QHBoxLayout()
        bottom.setSpacing(10)

        btn_cancel = QPushButton("취소")
        btn_cancel.setObjectName("GhostBtn")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.clicked.connect(self.reject)
        bottom.addWidget(btn_cancel)

        bottom.addStretch(1)

        btn_save = QPushButton("등록")
        btn_save.setObjectName("PrimaryBtn")
        btn_save.setCursor(Qt.PointingHandCursor)
        btn_save.clicked.connect(self._on_save)
        bottom.addWidget(btn_save)

        root.addLayout(bottom)

        self._apply_marks()

    def _apply_marks(self) -> None:
        for no, btn in self._no_buttons.items():
            v = self._answers.get(no, None)
            mark = "" if v is None else ("O" if v is True else "X")
            btn.setProperty("mark", mark)
            btn.style().unpolish(btn)
            btn.style().polish(btn)
            btn.update()

    def _toggle_no(self, no: int) -> None:
        v = self._answers.get(no, None)
        # cycle: None -> O -> X -> None
        if v is None:
            self._answers[no] = True
        elif v is True:
            self._answers[no] = False
        else:
            self._answers[no] = None
        self._apply_marks()
        self._sync_summary()

    def _set_all(self, val: bool) -> None:
        for no in list(self._answers.keys()):
            self._answers[no] = bool(val)
        self._apply_marks()
        self._sync_summary()

    def _clear_all(self) -> None:
        for no in list(self._answers.keys()):
            self._answers[no] = None
        self._apply_marks()
        self._sync_summary()

    def _sync_summary(self) -> None:
        total = len(self._answers)
        answered = sum(1 for _, v in self._answers.items() if v is not None)
        correct = sum(1 for _, v in self._answers.items() if v is True)
        if self.lbl_summary is not None:
            self.lbl_summary.setText(f"선택 {answered}/{total} · 정답 {correct}/{total}")

    def _on_save(self) -> None:
        total = len(self._answers)
        unanswered = [no for no, v in self._answers.items() if v is None]
        if unanswered:
            show_warning(self, "채점", f"아직 선택되지 않은 문항이 있습니다.\n\n미선택: {len(unanswered)}개")
            return
        if total <= 0:
            show_warning(self, "채점", "채점할 문항이 없습니다.")
            return
        self.accept()


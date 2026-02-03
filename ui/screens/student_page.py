"""
학생별 관리 페이지(수업 탭 내부)

요구사항:
- 학생별로 [학습지, 오답노트, 보고서] 3개 탭 제공
- 학생별 '학습지' 탭은 수업준비-학습지(WorksheetListScreen) 디자인을 그대로 사용
  - 단, "채점" 버튼만 1개 추가
- 수업준비-학습지 화면 코드는 절대 수정하지 않음(재사용/확장만)
"""

from __future__ import annotations

import os
import shutil
import tempfile
from datetime import datetime
from typing import Dict, List, Optional, Set, Tuple

from PyQt5.QtCore import Qt, QDate, pyqtSignal, QRect, QRectF
from PyQt5.QtGui import (
    QBrush,
    QColor,
    QConicalGradient,
    QFont,
    QPalette,
    QPainter,
    QPainterPath,
    QPixmap,
)
from PyQt5.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QStackedWidget,
    QScrollArea,
    QFrame,
    QGridLayout,
    QLineEdit,
    QTextEdit,
    QDateEdit,
    QProgressBar,
    QGraphicsDropShadowEffect,
    QSizePolicy,
    QDialog,
    QFileDialog,
    QApplication,
)

from core.models import Worksheet, SavedReport
from database.repositories import ProblemRepository, ReportRepository, WorksheetAssignmentRepository, WorksheetRepository
from services.report.report_service import aggregate_report
from ui.components.grading_dialog import GradingDialog
from ui.components.standard_message import show_info, show_warning
from ui.screens.worksheet_list import WorksheetListScreen, CompactStudySheetRow, StudySheetItem
from services.worksheet.hwp_composer import WorksheetHwpComposer, WorksheetComposeError
from processors.hwp.hwp_reader import HWPReader, HWPNotInstalledError, HWPInitializationError
from ui.components.standard_action_dialog import DialogAction, StandardActionDialog

try:
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg
    from matplotlib.figure import Figure
    _HAS_MATPLOTLIB = True
except ImportError:
    _HAS_MATPLOTLIB = False

try:
    from PyQt5.QtPrintSupport import QPrinter
    _HAS_QPRINTER = True
except ImportError:
    _HAS_QPRINTER = False


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


def _extract_problem_id(d: dict) -> str:
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


def _build_no_to_pid(numbered: list) -> Dict[int, str]:
    out: Dict[int, str] = {}
    for it in (numbered or []):
        if not isinstance(it, dict):
            continue
        try:
            no = int(it.get("no"))
        except Exception:
            continue
        pid = _extract_problem_id(it)
        if pid:
            out[no] = pid
    return out


def _safe_filename(name: str, *, fallback: str = "worksheet") -> str:
    s = (name or "").strip() or fallback
    bad = '<>:"/\\|?*'
    for ch in bad:
        s = s.replace(ch, "_")
    s = " ".join(s.split())
    return s[:120]


def _project_root() -> str:
    """프로젝트 루트(폴더 옮김 시 DB·오답노트 캐시가 함께 따라가도록)."""
    path = os.path.abspath(__file__)
    for _ in range(3):  # ui/screens/student_page.py -> 프로젝트 루트
        path = os.path.dirname(path)
    return path


def _ascii_temp_dir(sub: str) -> str:
    base = os.environ.get("SystemDrive", "C:") + os.sep
    d = os.path.join(base, "CH_LMS_TMP", sub)
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def _wrongnote_cache_dir() -> str:
    """오답노트 HWP 캐시 루트(프로젝트 db/wrongnote_cache → 폴더 옮김 시 함께 이동)."""
    root = _project_root()
    d = os.path.join(root, "db", "wrongnote_cache")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def _wrongnote_cached_hwp_path(student_id: str, worksheet_id: str) -> str:
    """해당 (학생, 학습지) 오답노트 HWP 캐시 파일 경로."""
    base = _wrongnote_cache_dir()
    sid = (student_id or "").strip() or "unknown"
    wid = (worksheet_id or "").strip() or "unknown"
    for ch in '<>:"/\\|?*':
        sid = sid.replace(ch, "_")
        wid = wid.replace(ch, "_")
    sub = os.path.join(base, sid[:64])
    try:
        os.makedirs(sub, exist_ok=True)
    except Exception:
        pass
    return os.path.join(sub, f"{wid[:64]}.hwp")


class StudentStudySheetRow(CompactStudySheetRow):
    """수업준비-학습지 Row 디자인을 그대로 쓰되, 채점/오답노트 버튼만 추가."""

    grade_requested = pyqtSignal(str)  # worksheet_id
    wrongnote_requested = pyqtSignal(str)  # worksheet_id

    def __init__(self, item: StudySheetItem, selected: bool, parent: Optional[QWidget] = None):
        super().__init__(item, selected=selected, parent=parent)

        # 맞은개수(채점 시에만 값 표시, 미채점이면 비움)
        self.score_badge = QLabel("")
        self.score_badge.setObjectName("ScoreBadge")
        self.score_badge.setFont(_font(8, bold=True))
        self.score_badge.setFixedSize(44, 26)
        self.score_badge.setAlignment(Qt.AlignCenter)
        self.score_badge.setStyleSheet(
            """
            QLabel#ScoreBadge {
                background: #F1F5F9;
                color: #0F172A;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 2px 6px;
            }
            """
        )
        is_graded = bool(getattr(item, "is_graded", False))
        summary = (getattr(item, "graded_summary", "") or "").strip() if is_graded else ""
        self.score_badge.setText(summary)
        self.score_badge.setToolTip("맞은 개수 / 전체" if summary else "채점 후 표시")

        self.btn_grade = QPushButton("채점")
        self.btn_grade.setObjectName("GradeBtn")
        self.btn_grade.setFocusPolicy(Qt.NoFocus)
        self.btn_grade.setCursor(Qt.PointingHandCursor)
        self.btn_grade.setFixedSize(56, 26)
        self.btn_grade.setFont(_font(8, bold=True))
        self.btn_grade.clicked.connect(lambda _=False: self.grade_requested.emit(self.item.id))

        wrongnote_enabled = bool(getattr(item, "wrongnote_enabled", False))
        self.btn_wrong = QPushButton("오답노트" if wrongnote_enabled else "오답노트 생성")
        self.btn_wrong.setObjectName("WrongBtn")
        self.btn_wrong.setFocusPolicy(Qt.NoFocus)
        self.btn_wrong.setCursor(Qt.PointingHandCursor)
        self.btn_wrong.setFixedSize(82, 26)
        self.btn_wrong.setFont(_font(8, bold=True))
        self.btn_wrong.clicked.connect(lambda _=False: self.wrongnote_requested.emit(self.item.id))
        if not is_graded:
            self.btn_wrong.setEnabled(False)
            self.btn_wrong.setToolTip("채점 완료 후 이용 가능")
        else:
            self.btn_wrong.setToolTip("오답노트 보기" if wrongnote_enabled else "오답노트 생성")

        # CompactRow 레이아웃의 마지막 2개(PDF/HWP) 앞에 삽입
        lay = self.layout()
        if lay is not None:
            try:
                idx = max(0, lay.count() - 2)
                # [점수] [채점] [오답노트] [PDF] [HWP]
                lay.insertWidget(idx, self.score_badge, alignment=Qt.AlignVCenter)
                lay.insertWidget(idx + 1, self.btn_grade, alignment=Qt.AlignVCenter)
                lay.insertWidget(idx + 2, self.btn_wrong, alignment=Qt.AlignVCenter)
            except Exception:
                # 폴백: 그냥 끝에 추가
                try:
                    lay.addWidget(self.score_badge, alignment=Qt.AlignVCenter)
                    lay.addWidget(self.btn_grade, alignment=Qt.AlignVCenter)
                    lay.addWidget(self.btn_wrong, alignment=Qt.AlignVCenter)
                except Exception:
                    pass

        # 버튼 스타일은 Row에만 로컬로 적용(수업준비 화면 영향 없음)
        self.btn_grade.setStyleSheet(
            """
            QPushButton#GradeBtn {
                padding: 0px;
                margin: 0px;
                background: #2563EB;
                border: 1.5px solid #1D4ED8;
                color: #FFFFFF;
                font-weight: 800;
                font-size: 9pt;
                border-radius: 6px;
            }
            QPushButton#GradeBtn:hover {
                background: #1D4ED8;
                border-color: #1E40AF;
            }
            """
        )

        self.btn_wrong.setStyleSheet(
            """
            QPushButton#WrongBtn {
                padding: 0px;
                margin: 0px;
                background: #EEF2FF;
                border: 1.5px solid #C7D2FE;
                color: #3730A3;
                font-weight: 800;
                font-size: 9pt;
                border-radius: 6px;
            }
            QPushButton#WrongBtn:hover:enabled {
                background: #E0E7FF;
                border-color: #A5B4FC;
                color: #312E81;
            }
            QPushButton#WrongBtn:disabled {
                background: #F1F5F9;
                border: 1.5px solid #E2E8F0;
                color: #94A3B8;
            }
            """
        )


class StudentWorksheetListScreen(WorksheetListScreen):
    grading_saved = pyqtSignal(str)  # worksheet_id
    wrongnote_ready = pyqtSignal(str)  # worksheet_id
    """학생 페이지용 학습지 리스트(출제된 것만 로드 + 채점/오답노트 버튼)."""

    def __init__(
        self,
        db_connection,
        *,
        student_id: str,
        student_name: str,
        student_grade: str,
        parent: Optional[QWidget] = None,
    ):
        self.student_id = (student_id or "").strip()
        self.student_name = (student_name or "").strip()
        self.student_grade = (student_grade or "").strip()
        super().__init__(db_connection, parent=parent)

        # 학생 페이지에서는 "학습지 생성" 버튼을 숨김(수업준비 전용 기능)
        try:
            btn_create = self.findChild(QPushButton, "create")
            if btn_create is not None:
                btn_create.hide()
        except Exception:
            pass

        # 학생 페이지에서는 출제/삭제 버튼은 숨김(목록/다운로드/채점 중심)
        try:
            if getattr(self, "btn_assign", None) is not None:
                self.btn_assign.hide()
            if getattr(self, "btn_bulk_delete", None) is not None:
                self.btn_bulk_delete.hide()
        except Exception:
            pass

        # 목록 로드는 "출제된 것만"으로 오버라이드
        self.reload_from_db()

    def reload_from_db(self) -> None:
        """worksheet_assignments에서 학생에게 출제된 학습지만 로드."""
        self._items = []

        if not self.ws_repo or not self.db_connection:
            self.refresh_list()
            return
        if not self.db_connection.is_connected():
            self._selected_ids.clear()
            self.refresh_list()
            return
        if not self.student_id:
            self._selected_ids.clear()
            self.refresh_list()
            return

        try:
            assign_repo = WorksheetAssignmentRepository(self.db_connection)
            assigns = assign_repo.list_for_student(self.student_id)
        except Exception:
            assigns = []

        ws_ids = [str(a.get("worksheet_id") or "").strip() for a in assigns if str(a.get("worksheet_id") or "").strip()]
        if not ws_ids:
            self._selected_ids.clear()
            self.refresh_list()
            return

        worksheets = self.ws_repo.list_by_ids(ws_ids)
        by_id: Dict[str, Worksheet] = {str(w.id): w for w in worksheets if w and w.id}

        items = []
        for a in assigns:
            wid = str(a.get("worksheet_id") or "").strip()
            ws = by_id.get(wid)
            if not ws:
                continue
            dt = ws.created_at
            date_str = dt.strftime("%Y.%m.%d") if dt else ""
            it = StudySheetItem(
                id=wid,
                grade=(ws.grade or "").strip(),
                type_text=(ws.type_text or "").strip(),
                title=(ws.title or "").strip(),
                date=date_str,
                teacher=(ws.creator or "").strip(),
                has_hwp=bool(getattr(ws, "hwp_file_id", None)),
                has_pdf=bool(getattr(ws, "pdf_file_id", None)),
            )
            # 채점 결과 표시용(동적 속성)
            total_q = int(a.get("total_questions") or 0)
            correct = int(a.get("correct_count") or 0)
            is_graded = str(a.get("status") or "") == "graded" and total_q > 0
            setattr(it, "is_graded", bool(is_graded))
            wc = a.get("wrong_count")
            if wc is None and is_graded:
                # 레거시: wrong_count가 없으면 점수로부터 계산(오답노트 생성 시에도 backfill 됨)
                wc = max(0, int(total_q) - int(correct))
            setattr(it, "wrong_count", int(wc or 0))
            setattr(it, "wrongnote_enabled", bool(a.get("wrongnote_enabled", False)))
            if is_graded:
                setattr(it, "graded_summary", f"{correct}/{total_q}")
            items.append(it)

        # 수업준비-학습지와 동일: 학습지 생성일(created_at) 최신순
        items.sort(
            key=lambda it: (getattr(by_id.get(it.id), "created_at", None) or datetime.min),
            reverse=True,
        )
        self._items = items
        alive = {it.id for it in self._items}
        self._selected_ids = {sid for sid in self._selected_ids if sid in alive}
        self.refresh_list()

    def refresh_list(self) -> None:
        # super와 동일 로직이지만 Row만 교체(채점 버튼 추가)
        while self.list_layout.count() > 1:
            it = self.list_layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.setParent(None)

        query = (self.search_input.text() or "").strip().lower()

        visible_items = []
        for it in self._items:
            if not self._grade_match(it.grade):
                continue
            if query:
                hay = f"{it.grade} {it.type_text} {it.title} {it.date} {it.teacher}".lower()
                if query not in hay:
                    continue
            visible_items.append(it)

        self._visible_ids = [it.id for it in visible_items]

        if not visible_items:
            empty = QLabel("표시할 학습지가 없습니다.")
            empty.setFont(_font(10, bold=True))
            empty.setStyleSheet("color: #475569;")
            empty.setAlignment(Qt.AlignCenter)
            empty.setMinimumHeight(160)
            self.list_layout.insertWidget(0, empty)
            self._sync_action_bar_state()
            return

        for it in visible_items:
            row = StudentStudySheetRow(it, selected=it.id in self._selected_ids)
            row.selected_changed.connect(self._on_row_selected_changed)
            row.download_requested.connect(self._on_row_download_requested)
            row.grade_requested.connect(self._on_grade_requested)
            row.wrongnote_requested.connect(self._on_wrongnote_requested)
            self.list_layout.insertWidget(self.list_layout.count() - 1, row)

        self._sync_action_bar_state()

    def _on_grade_requested(self, worksheet_id: str) -> None:
        if not self.db_connection or not self.db_connection.is_connected():
            show_warning(self, "채점", "DB에 연결할 수 없습니다.")
            return
        if not self.student_id:
            show_warning(self, "채점", "학생 정보(student_id)가 없습니다.")
            return

        ws_repo = WorksheetRepository(self.db_connection)
        ws = ws_repo.find_by_id(worksheet_id)
        if not ws:
            show_warning(self, "채점", "학습지 정보를 찾을 수 없습니다.")
            return

        # 문항 목록(번호-문제ID)
        numbered = list(ws.numbered or [])
        if not numbered:
            numbered = [{"no": i + 1, "problem_id": pid} for i, pid in enumerate(list(ws.problem_ids or []))]

        # 문제 로드(단원 통계)
        pids = [str(x.get("problem_id") or "").strip() for x in numbered if str(x.get("problem_id") or "").strip()]
        probs_by_id: Dict[str, object] = {}
        try:
            pr = ProblemRepository(self.db_connection)
            probs = pr.list_by_ids(pids)
            probs_by_id = {str(p.id): p for p in probs if p and p.id}
        except Exception:
            probs_by_id = {}

        # 기존 채점 결과(있으면 prefill)
        existing: Dict[int, bool] = {}
        try:
            ar = WorksheetAssignmentRepository(self.db_connection)
            doc = ar.find_one(worksheet_id=worksheet_id, student_id=self.student_id)
            if doc and isinstance(doc.get("answers"), list):
                for a in doc.get("answers") or []:
                    try:
                        no = int(a.get("no"))
                        existing[no] = bool(a.get("is_correct"))
                    except Exception:
                        continue
        except Exception:
            existing = {}

        dlg = GradingDialog(
            worksheet=ws,
            numbered=numbered,
            problems_by_id=probs_by_id,  # type: ignore[arg-type]
            existing_answers=existing,
            parent=self,
        )
        if dlg.exec_() != dlg.Accepted:
            return

        payload = dlg.result_payload()
        total_q = int(payload.get("total_questions") or 0)
        correct = int(payload.get("correct_count") or 0)
        answers = payload.get("answers") or []
        unit_stats = payload.get("unit_stats") or {}

        ok = False
        try:
            ar = WorksheetAssignmentRepository(self.db_connection)
            ok = ar.save_grading(
                worksheet_id=worksheet_id,
                student_id=self.student_id,
                total_questions=total_q,
                correct_count=correct,
                answers=answers,
                unit_stats=unit_stats,
                assigned_by="",
            )
        except Exception:
            ok = False

        if not ok:
            show_warning(self, "채점", "채점 결과 저장에 실패했습니다.")
            return

        show_info(self, "채점 완료", f"{total_q}개 중 {correct}개 정답으로 저장했습니다.")
        self.reload_from_db()
        self.grading_saved.emit(worksheet_id)

    def _on_wrongnote_requested(self, worksheet_id: str) -> None:
        if not self.db_connection or not self.db_connection.is_connected():
            show_warning(self, "오답노트", "DB에 연결할 수 없습니다.")
            return
        if not self.student_id:
            show_warning(self, "오답노트", "학생 정보(student_id)가 없습니다.")
            return

        # 출제 문서 확인(채점 여부/오답 여부)
        ar = WorksheetAssignmentRepository(self.db_connection)
        doc = ar.find_one(worksheet_id=worksheet_id, student_id=self.student_id)
        if not doc:
            show_warning(self, "오답노트", "출제 정보를 찾을 수 없습니다.")
            return
        if str(doc.get("status") or "") != "graded":
            show_warning(self, "오답노트", "채점이 완료된 학습지에서만 오답노트를 생성할 수 있습니다.")
            return

        wrong_ids = [str(x).strip() for x in (doc.get("wrong_problem_ids") or []) if str(x).strip()]

        # ✅ 레거시 채점 데이터 보정:
        # 과거 채점 기록에는 wrong_problem_ids가 없을 수 있으므로, answers + worksheet.numbered로 복구합니다.
        if not wrong_ids:
            ws_repo = WorksheetRepository(self.db_connection)
            ws = ws_repo.find_by_id(worksheet_id)
            if ws:
                no_to_pid = _build_no_to_pid(list(ws.numbered or []))
                # numbered가 비어있으면 problem_ids로 매핑
                if not no_to_pid and getattr(ws, "problem_ids", None):
                    try:
                        for i, pid in enumerate(list(ws.problem_ids or []), start=1):
                            pid2 = str(pid or "").strip()
                            if pid2:
                                no_to_pid[i] = pid2
                    except Exception:
                        pass

                recovered: List[str] = []
                for a in (doc.get("answers") or []):
                    if not isinstance(a, dict):
                        continue
                    try:
                        is_correct = a.get("is_correct")
                        # is_correct가 문자열로 저장된 레거시도 방어
                        if isinstance(is_correct, str):
                            ic = is_correct.strip().lower()
                            is_correct = ic in ("true", "1", "o", "ok", "yes")
                        if bool(is_correct):
                            continue  # 정답이면 스킵
                        no = int(a.get("no"))
                    except Exception:
                        continue
                    pid = str(a.get("problem_id") or "").strip()
                    if not pid:
                        pid = str(no_to_pid.get(no, "") or "").strip()
                    if pid:
                        recovered.append(pid)

                # 중복 제거(순서 유지)
                seen: Set[str] = set()
                wrong_ids = []
                for pid in recovered:
                    if pid in seen:
                        continue
                    seen.add(pid)
                    wrong_ids.append(pid)

                if wrong_ids:
                    try:
                        ar.set_wrong_info(worksheet_id=worksheet_id, student_id=self.student_id, wrong_problem_ids=wrong_ids)
                    except Exception:
                        pass

        if not wrong_ids:
            show_info(self, "오답노트", "틀린 문항이 없습니다. (오답노트 생성 불필요)")
            return

        ws_repo = WorksheetRepository(self.db_connection)
        ws = ws_repo.find_by_id(worksheet_id)
        if not ws:
            show_warning(self, "오답노트", "학습지 정보를 찾을 수 없습니다.")
            return

        title = f"{(ws.title or '').strip()}-{self.student_name}-오답"
        date_str = ""
        try:
            dt = ws.created_at
            date_str = dt.strftime("%Y.%m.%d") if dt else ""
        except Exception:
            date_str = ""
        if not date_str:
            date_str = datetime.now().strftime("%Y.%m.%d")

        cache_path = _wrongnote_cached_hwp_path(self.student_id, worksheet_id)
        already_enabled = bool(doc.get("wrongnote_enabled", False))

        # 이미 활성화되어 있고 캐시 HWP가 있으면 탭만 이동
        if already_enabled and os.path.isfile(cache_path):
            self.wrongnote_ready.emit(worksheet_id)
            return

        # HWP 재조합 후 캐시에 저장
        composer = WorksheetHwpComposer(self.db_connection)
        try:
            composer.compose(
                problem_ids=list(wrong_ids),
                output_path=cache_path,
                title=title,
                teacher=(ws.creator or "").strip(),
                date_str=date_str,
            )
        except WorksheetComposeError as e:
            show_warning(self, "오답노트", f"HWP 생성에 실패했습니다.\n\n{e}")
            return
        except Exception as e:
            show_warning(self, "오답노트", f"HWP 생성 중 오류가 발생했습니다.\n\n{e}")
            return

        if not already_enabled:
            ok = ar.enable_wrongnote(worksheet_id=worksheet_id, student_id=self.student_id, title=title)
            if not ok:
                show_warning(self, "오답노트", "오답노트 활성화에 실패했습니다.")
                return

        if getattr(composer, "_template_missing", False):
            show_info(
                self,
                "오답노트",
                "오답노트를 생성했습니다. 오답노트 탭에서 HWP/PDF를 다운로드할 수 있습니다.\n\n"
                "참고: 이 PC에서 템플릿 파일을 찾지 못해 빈 문서로 생성되었습니다. "
                "학습지/오답노트 서식을 쓰려면 exe가 있는 폴더에 templates 폴더(worksheet_template.hwp 포함)를 두세요.",
            )
        else:
            show_info(self, "오답노트", "오답노트를 생성했습니다. 오답노트 탭에서 HWP/PDF를 다운로드할 수 있습니다.")
        self.reload_from_db()
        self.wrongnote_ready.emit(worksheet_id)


class StudentWrongNoteListScreen(WorksheetListScreen):
    """학생 페이지용 오답노트 탭: '오답노트 생성'된 항목만 표시(틀린 문항 기반)."""

    def __init__(
        self,
        db_connection,
        *,
        student_id: str,
        student_name: str,
        parent: Optional[QWidget] = None,
    ):
        self.student_id = (student_id or "").strip()
        self.student_name = (student_name or "").strip()
        super().__init__(db_connection, parent=parent)

        # 오답노트 탭에서는 상단 생성/출제/삭제는 숨김
        try:
            btn_create = self.findChild(QPushButton, "create")
            if btn_create is not None:
                btn_create.hide()
        except Exception:
            pass
        try:
            if getattr(self, "btn_assign", None) is not None:
                self.btn_assign.hide()
            if getattr(self, "btn_bulk_delete", None) is not None:
                self.btn_bulk_delete.hide()
        except Exception:
            pass

        # 체크박스 기반 일괄 기능은 당장은 의미가 없으므로 숨김(디자인 유지)
        try:
            if getattr(self, "chk_select_all", None) is not None:
                self.chk_select_all.hide()
            if getattr(self, "lbl_selected", None) is not None:
                self.lbl_selected.hide()
            if getattr(self, "btn_bulk_download", None) is not None:
                self.btn_bulk_download.hide()
        except Exception:
            pass

        self.reload_from_db()

    def reload_from_db(self) -> None:
        self._items = []
        if not self.ws_repo or not self.db_connection or not self.db_connection.is_connected() or not self.student_id:
            self.refresh_list()
            return

        ar = WorksheetAssignmentRepository(self.db_connection)
        assigns = ar.list_wrongnotes_for_student(self.student_id)
        ws_ids = [str(a.get("worksheet_id") or "").strip() for a in assigns if str(a.get("worksheet_id") or "").strip()]
        if not ws_ids:
            self.refresh_list()
            return

        worksheets = self.ws_repo.list_by_ids(ws_ids)
        by_id: Dict[str, Worksheet] = {str(w.id): w for w in worksheets if w and w.id}

        items = []
        for a in assigns:
            wid = str(a.get("worksheet_id") or "").strip()
            ws = by_id.get(wid)
            if not ws:
                continue
            dt = ws.created_at
            date_str = dt.strftime("%Y.%m.%d") if dt else ""
            title = (a.get("wrongnote_title") or "").strip() or f"{(ws.title or '').strip()}-{self.student_name}-오답"

            it = StudySheetItem(
                id=wid,  # 원본 worksheet_id를 유지(추후 연결/갱신 용이)
                grade=(ws.grade or "").strip(),
                type_text=(ws.type_text or "").strip(),  # 출처(유형) 유지
                title=title,
                date=date_str,
                teacher=(ws.creator or "").strip(),
                # ✅ 오답노트는 템플릿 기반으로 "생성 가능한 HWP"가 있으므로 버튼을 활성화
                has_hwp=True,
                has_pdf=False,
            )

            # 필요한 메타(동적 속성)
            setattr(it, "wrong_problem_ids", list(a.get("wrong_problem_ids") or []))
            setattr(it, "wrongnote_title", title)
            setattr(it, "source_teacher", (ws.creator or "").strip())
            setattr(it, "source_created_at", ws.created_at)

            # 오답노트 채점 결과 표시(있으면)
            wn_t = int(a.get("wrongnote_total_questions") or 0)
            wn_c = int(a.get("wrongnote_correct_count") or 0)
            if str(a.get("wrongnote_status") or "") == "graded" and wn_t > 0:
                setattr(it, "graded_summary", f"{wn_c}/{wn_t}")
                setattr(it, "is_graded", True)
            else:
                setattr(it, "is_graded", False)
                # 오답 개수 배지(있으면)
                wc = int(a.get("wrong_count") or 0)
                if wc > 0:
                    setattr(it, "graded_summary", f"오답 {wc}")

            wc = int(a.get("wrong_count") or 0)
            items.append(it)

        # 수업준비-학습지와 동일: 학습지 생성일(created_at) 최신순
        items.sort(
            key=lambda it: (getattr(by_id.get(it.id), "created_at", None) or datetime.min),
            reverse=True,
        )
        self._items = items
        self.refresh_list()

    def refresh_list(self) -> None:
        # 오답노트 탭: Row에 "채점" 버튼을 추가로 붙임
        while self.list_layout.count() > 1:
            it = self.list_layout.takeAt(0)
            w = it.widget()
            if w is not None:
                w.setParent(None)

        query = (self.search_input.text() or "").strip().lower()
        visible_items = []
        for it in self._items:
            if not self._grade_match(it.grade):
                continue
            if query:
                hay = f"{it.grade} {it.type_text} {it.title} {it.date} {it.teacher}".lower()
                if query not in hay:
                    continue
            visible_items.append(it)

        self._visible_ids = [it.id for it in visible_items]

        if not visible_items:
            empty = QLabel("표시할 오답노트가 없습니다.")
            empty.setFont(_font(10, bold=True))
            empty.setStyleSheet("color: #475569;")
            empty.setAlignment(Qt.AlignCenter)
            empty.setMinimumHeight(160)
            self.list_layout.insertWidget(0, empty)
            self._sync_action_bar_state()
            return

        for it in visible_items:
            row = WrongNoteRow(it, selected=it.id in self._selected_ids)
            row.selected_changed.connect(self._on_row_selected_changed)
            row.download_requested.connect(self._on_wrongnote_download_requested)
            row.grade_requested.connect(self._on_wrongnote_grade_requested)
            self.list_layout.insertWidget(self.list_layout.count() - 1, row)

        self._sync_action_bar_state()

    def _on_wrongnote_grade_requested(self, worksheet_id: str) -> None:
        if not self.db_connection or not self.db_connection.is_connected():
            show_warning(self, "채점", "DB에 연결할 수 없습니다.")
            return
        if not self.student_id:
            show_warning(self, "채점", "학생 정보(student_id)가 없습니다.")
            return

        ar = WorksheetAssignmentRepository(self.db_connection)
        doc = ar.find_one(worksheet_id=worksheet_id, student_id=self.student_id) or {}
        wrong_ids = [str(x).strip() for x in (doc.get("wrong_problem_ids") or []) if str(x).strip()]
        if not wrong_ids:
            show_warning(self, "채점", "오답노트 문항이 없습니다.")
            return

        ws_repo = WorksheetRepository(self.db_connection)
        base_ws = ws_repo.find_by_id(worksheet_id)
        if not base_ws:
            show_warning(self, "채점", "학습지 정보를 찾을 수 없습니다.")
            return

        title = (doc.get("wrongnote_title") or "").strip() or f"{(base_ws.title or '').strip()}-{self.student_name}-오답"

        numbered = [{"no": i + 1, "problem_id": pid} for i, pid in enumerate(wrong_ids)]
        # 문제 로드(단원 통계)
        probs_by_id: Dict[str, object] = {}
        try:
            pr = ProblemRepository(self.db_connection)
            probs = pr.list_by_ids(list(wrong_ids))
            probs_by_id = {str(p.id): p for p in probs if p and p.id}
        except Exception:
            probs_by_id = {}

        # 기존 오답노트 채점 결과(있으면 prefill)
        existing: Dict[int, bool] = {}
        if isinstance(doc.get("wrongnote_answers"), list):
            for a in doc.get("wrongnote_answers") or []:
                try:
                    no = int(a.get("no"))
                    existing[no] = bool(a.get("is_correct"))
                except Exception:
                    continue

        fake_ws = Worksheet(
            id=base_ws.id,
            title=title,
            grade=base_ws.grade,
            type_text=base_ws.type_text,
            creator=base_ws.creator,
            created_at=base_ws.created_at,
            problem_ids=list(wrong_ids),
            numbered=numbered,
        )

        dlg = GradingDialog(
            worksheet=fake_ws,
            numbered=numbered,
            problems_by_id=probs_by_id,  # type: ignore[arg-type]
            existing_answers=existing,
            parent=self,
        )
        if dlg.exec_() != dlg.Accepted:
            return

        payload = dlg.result_payload()
        total_q = int(payload.get("total_questions") or 0)
        correct = int(payload.get("correct_count") or 0)
        answers = payload.get("answers") or []
        unit_stats = payload.get("unit_stats") or {}

        ok = False
        try:
            ok = ar.save_wrongnote_grading(
                worksheet_id=worksheet_id,
                student_id=self.student_id,
                total_questions=total_q,
                correct_count=correct,
                answers=answers,
                unit_stats=unit_stats,
                assigned_by="",
            )
        except Exception:
            ok = False

        if not ok:
            show_warning(self, "채점", "오답노트 채점 결과 저장에 실패했습니다.")
            return

        show_info(self, "채점 완료", f"{total_q}개 중 {correct}개 정답으로 저장했습니다.")
        self.reload_from_db()

    def _on_wrongnote_download_requested(self, worksheet_id: str, kind: str) -> None:
        """
        오답노트 HWP/PDF: 학습지 탭에서 생성된 캐시 HWP만 사용하여 저장/열기(재생성 없음).
        """
        if not self.db_connection or not self.db_connection.is_connected():
            show_warning(self, "다운로드", "DB에 연결할 수 없습니다.")
            return
        if not self.student_id:
            show_warning(self, "다운로드", "학생 정보(student_id)가 없습니다.")
            return

        cached_hwp = _wrongnote_cached_hwp_path(self.student_id, worksheet_id)
        if not os.path.isfile(cached_hwp):
            show_warning(
                self,
                "다운로드",
                "오답노트 HWP가 생성되지 않았습니다.\n학습지 탭에서 해당 학습지의 '오답노트 생성'을 눌러 주세요.",
            )
            return

        ar = WorksheetAssignmentRepository(self.db_connection)
        doc = ar.find_one(worksheet_id=worksheet_id, student_id=self.student_id) or {}
        ws_repo = WorksheetRepository(self.db_connection)
        ws = ws_repo.find_by_id(worksheet_id)
        title = (doc.get("wrongnote_title") or "").strip()
        if ws:
            title = title or f"{(ws.title or '').strip()}-{self.student_name}-오답"
        title = title or "오답노트"
        safe = _safe_filename(title, fallback="wrongnote")

        k = (kind or "").strip().upper()
        if k not in ("HWP", "PDF"):
            show_warning(self, "다운로드", f"지원하지 않는 형식입니다: {kind}")
            return

        dlg = StandardActionDialog(
            parent=self,
            title=k,
            message=f"{k} 작업을 선택하세요.",
            actions=[
                DialogAction(key="save", label="저장", is_primary=False),
                DialogAction(key="open", label="열기", is_primary=True),
                DialogAction(key="cancel", label="취소", is_primary=False),
            ],
            min_width=360,
        )
        dlg.exec_()
        if dlg.selected_key not in ("save", "open"):
            return

        # HWP: 캐시 파일 복사 또는 열기
        if k == "HWP":
            if dlg.selected_key == "save":
                default_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{safe}.hwp")
                out, _ = QFileDialog.getSaveFileName(self, "HWP 저장", default_path, "HWP Files (*.hwp)")
                if not out:
                    return
                try:
                    shutil.copy2(cached_hwp, out)
                    show_info(self, "완료", "HWP가 저장되었습니다.")
                except Exception as e:
                    show_warning(self, "HWP 저장 실패", str(e))
                return
            try:
                os.startfile(cached_hwp)  # type: ignore[attr-defined]
            except Exception as e:
                show_warning(self, "열기 실패", f"HWP를 열 수 없습니다.\n\n{e}")
            return

        # PDF: 캐시 HWP를 HWPReader로 열어 PDF 내보내기
        if k == "PDF":
            try:
                if dlg.selected_key == "save":
                    default_path = os.path.join(os.path.expanduser("~"), "Desktop", f"{safe}.pdf")
                    out, _ = QFileDialog.getSaveFileName(self, "PDF 저장", default_path, "PDF Files (*.pdf)")
                    if not out:
                        return
                    with HWPReader() as reader:
                        opened = reader.open_document(cached_hwp)
                        if not opened:
                            raise RuntimeError("한글(HWP)에서 파일을 열 수 없습니다.")
                        ok = reader.export_pdf(out)
                        try:
                            reader.close_document()
                        except Exception:
                            pass
                        if not ok:
                            raise RuntimeError("PDF로 내보내기에 실패했습니다.")
                    show_info(self, "완료", "PDF가 저장되었습니다.")
                    return

                out_pdf = os.path.join(_ascii_temp_dir("wrongnote_open"), f"{safe}.pdf")
                with HWPReader() as reader:
                    opened = reader.open_document(cached_hwp)
                    if not opened:
                        raise RuntimeError("한글(HWP)에서 파일을 열 수 없습니다.")
                    ok = reader.export_pdf(out_pdf)
                    try:
                        reader.close_document()
                    except Exception:
                        pass
                    if not ok:
                        raise RuntimeError("PDF로 내보내기에 실패했습니다.")
                try:
                    os.startfile(out_pdf)  # type: ignore[attr-defined]
                except Exception as e:
                    show_warning(self, "열기 실패", f"PDF를 열 수 없습니다.\n\n{e}")
            except (HWPNotInstalledError, HWPInitializationError) as e:
                show_warning(self, "PDF", str(e))
            except Exception as e:
                show_warning(self, "PDF", f"PDF 생성/저장에 실패했습니다.\n\n{e}")


class WrongNoteRow(CompactStudySheetRow):
    """오답노트 리스트 Row: '채점' 버튼을 추가(다운로드 버튼은 기본 유지)."""

    grade_requested = pyqtSignal(str)  # worksheet_id

    def __init__(self, item: StudySheetItem, selected: bool, parent: Optional[QWidget] = None):
        super().__init__(item, selected=selected, parent=parent)

        self.btn_grade = QPushButton("채점")
        self.btn_grade.setObjectName("GradeBtn")
        self.btn_grade.setFocusPolicy(Qt.NoFocus)
        self.btn_grade.setCursor(Qt.PointingHandCursor)
        self.btn_grade.setFixedSize(56, 26)
        self.btn_grade.setFont(_font(8, bold=True))
        self.btn_grade.clicked.connect(lambda _=False: self.grade_requested.emit(self.item.id))
        self.btn_grade.setStyleSheet(
            """
            QPushButton#GradeBtn {
                padding: 0px;
                margin: 0px;
                background: #2563EB;
                border: 1.5px solid #1D4ED8;
                color: #FFFFFF;
                font-weight: 800;
                font-size: 9pt;
                border-radius: 6px;
            }
            QPushButton#GradeBtn:hover {
                background: #1D4ED8;
                border-color: #1E40AF;
            }
            """
        )

        # 점수 배지(있으면 표시)
        self.score_badge = QLabel("")
        self.score_badge.setObjectName("ScoreBadge")
        self.score_badge.setFont(_font(8, bold=True))
        self.score_badge.setStyleSheet(
            """
            QLabel#ScoreBadge {
                background: #F1F5F9;
                color: #0F172A;
                border: 1px solid #E2E8F0;
                border-radius: 6px;
                padding: 2px 8px;
            }
            """
        )
        summary = getattr(item, "graded_summary", "") or ""
        if summary:
            self.score_badge.setText(summary)
            self.score_badge.show()
        else:
            self.score_badge.hide()

        lay = self.layout()
        if lay is not None:
            try:
                idx = max(0, lay.count() - 2)  # PDF/HWP 앞
                lay.insertWidget(idx, self.score_badge, alignment=Qt.AlignVCenter)
                lay.insertWidget(idx + 1, self.btn_grade, alignment=Qt.AlignVCenter)
            except Exception:
                try:
                    lay.addWidget(self.score_badge, alignment=Qt.AlignVCenter)
                    lay.addWidget(self.btn_grade, alignment=Qt.AlignVCenter)
                except Exception:
                    pass


# ----- 보고서 탭 -----


class _ReportListRow(QFrame):
    """저장된 보고서 목록 한 행. 1단 풀폭, 호버 시 행 전체 #E8F0FE, 내부 라벨 배경 투명. [제목/생성일/유형칩/수정/삭제]."""

    clicked = pyqtSignal(str)  # report_id
    edit_requested = pyqtSignal(str)  # report_id
    delete_requested = pyqtSignal(str)  # report_id

    def __init__(self, report: SavedReport, student_name: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.report_id = (report.id or "").strip()
        self.setCursor(Qt.PointingHandCursor)
        self.setObjectName("ReportRow")
        self.setStyleSheet(
            """
            QFrame#ReportRow {
                background-color: #FFFFFF;
                border: none;
                border-bottom: 1px solid #E5E7EB;
                padding: 0;
            }
            QFrame#ReportRow:hover {
                background-color: #E8F0FE;
            }
            QFrame#ReportRow QLabel {
                background-color: transparent;
                background: none;
                color: #222222;
                font-size: 11pt;
            }
            QFrame#ReportRow QLabel#ReportTypeChip {
                background-color: #E8F0FE;
                color: #007BFF;
                border-radius: 12px;
            }
            QPushButton#ReportRowEditBtn {
                background: #E0E7FF;
                color: #3730A3;
                border: none;
                border-radius: 6px;
                font-size: 11pt;
                font-weight: bold;
                padding: 4px 10px;
                min-width: 64px;
            }
            QPushButton#ReportRowEditBtn:hover {
                background: #C7D2FE;
            }
            QPushButton#ReportRowDelBtn {
                background: transparent;
                color: #FF4D4F;
                border: 1px solid #FF4D4F;
                border-radius: 4px;
                font-size: 11pt;
                font-weight: bold;
                padding: 5px 10px;
            }
            QPushButton#ReportRowDelBtn:hover {
                background: #FFF1F0;
            }
            """
        )
        lay = QHBoxLayout(self)
        lay.setContentsMargins(20, 12, 20, 12)
        lay.setSpacing(20)

        period = f"{(report.period_start or '').replace('-', '.')} ~ {(report.period_end or '').replace('-', '.')}"
        title = f"{(student_name or '학생').strip()}학생의 {period} 보고서"
        created = ""
        if report.created_at:
            try:
                created = report.created_at.strftime("%Y.%m.%d %H:%M")
            except Exception:
                created = str(report.created_at)

        lbl_title = QLabel(title)
        lbl_title.setFont(_font(12, bold=True))
        lbl_title.setStyleSheet("background: none; color: #222222; font-size: 12pt;")
        lbl_title.setWordWrap(True)
        lbl_created = QLabel(created)
        lbl_created.setFont(_font(11))
        lbl_created.setStyleSheet("background: none; color: #222222; font-size: 11pt;")
        lbl_created.setFixedWidth(120)
        lbl_created.setAlignment(Qt.AlignCenter)

        type_chip = QLabel("학습 보고서")
        type_chip.setObjectName("ReportTypeChip")
        type_chip.setFont(_font(10, bold=True))
        type_chip.setAlignment(Qt.AlignCenter)
        type_chip.setFixedWidth(90)
        type_chip.setFixedHeight(26)

        btn_edit = QPushButton("수정")
        btn_edit.setObjectName("ReportRowEditBtn")
        btn_edit.setFixedSize(72, 28)
        btn_edit.setFocusPolicy(Qt.NoFocus)
        btn_edit.setCursor(Qt.PointingHandCursor)
        btn_edit.clicked.connect(lambda: self.edit_requested.emit(self.report_id))
        btn_del = QPushButton("삭제")
        btn_del.setObjectName("ReportRowDelBtn")
        btn_del.setFixedSize(72, 28)
        btn_del.setFocusPolicy(Qt.NoFocus)
        btn_del.setCursor(Qt.PointingHandCursor)
        btn_del.clicked.connect(lambda: self.delete_requested.emit(self.report_id))

        lay.addWidget(lbl_title, 1)
        lay.addWidget(lbl_created)
        lay.addWidget(type_chip)
        lay.addWidget(btn_edit)
        lay.addWidget(btn_del)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.report_id)
        super().mousePressEvent(event)


class ReportCreateModal(QDialog):
    """보고서 생성 모달: 기간·코멘트 입력 + 해당 기간 미리보기 + 저장/취소."""

    def __init__(
        self,
        db_connection,
        student_id: str,
        student_name: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self._db = db_connection
        self._student_id = (student_id or "").strip()
        self._student_name = (student_name or "").strip()
        self._current_snapshot = None
        self.setWindowTitle("새 보고서 생성")
        self.setMinimumSize(920, 720)
        self.resize(960, 780)
        self.setStyleSheet("QDialog { background-color: #F8FAFC; }")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(16)

        row_dates = QHBoxLayout()
        row_dates.addWidget(QLabel("기간:"))
        self._date_start = QDateEdit()
        self._date_start.setCalendarPopup(True)
        self._date_start.setDate(QDate.currentDate().addMonths(-1))
        self._date_end = QDateEdit()
        self._date_end.setCalendarPopup(True)
        self._date_end.setDate(QDate.currentDate())
        row_dates.addWidget(self._date_start)
        row_dates.addWidget(QLabel(" ~ "))
        row_dates.addWidget(self._date_end)
        row_dates.addStretch(1)
        btn_preview = QPushButton("미리보기")
        btn_preview.setCursor(Qt.PointingHandCursor)
        btn_preview.setFixedHeight(32)
        btn_preview.setStyleSheet(
            "QPushButton { background: #475569; color: #FFF; border: none; border-radius: 6px; padding: 0 14px; }"
            "QPushButton:hover { background: #334155; }"
        )
        btn_preview.clicked.connect(self._on_preview)
        row_dates.addWidget(btn_preview)
        lay.addLayout(row_dates)

        lay.addWidget(QLabel("학습 코멘트 (선택):"))
        self._comment_input = QTextEdit()
        self._comment_input.setPlaceholderText("학습 코멘트를 입력하세요. 여러 줄 입력 가능합니다.")
        self._comment_input.setMinimumHeight(100)
        self._comment_input.setMaximumHeight(180)
        self._comment_input.setStyleSheet(
            "QTextEdit { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px; font-size: 11pt; }"
        )
        lay.addWidget(self._comment_input)

        lay.addWidget(QLabel("보고서 미리보기:"))
        self._preview_scroll = QScrollArea()
        self._preview_scroll.setWidgetResizable(True)
        self._preview_scroll.setFrameShape(QFrame.NoFrame)
        self._preview_scroll.setStyleSheet("QScrollArea { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; }")
        self._preview_placeholder = QLabel("기간을 선택한 뒤 '미리보기'를 누르면 해당 기간 보고서가 여기에 표시됩니다.")
        self._preview_placeholder.setStyleSheet("color: #94A3B8; padding: 40px;")
        self._preview_placeholder.setWordWrap(True)
        self._preview_placeholder.setAlignment(Qt.AlignCenter)
        self._preview_scroll.setWidget(self._preview_placeholder)
        lay.addWidget(self._preview_scroll, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setFixedHeight(36)
        btn_cancel.setStyleSheet(
            "QPushButton { background: #E2E8F0; color: #334155; border: none; border-radius: 8px; padding: 0 20px; }"
            "QPushButton:hover { background: #CBD5E1; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_save = QPushButton("저장")
        btn_save.setCursor(Qt.PointingHandCursor)
        btn_save.setFixedHeight(36)
        btn_save.setStyleSheet(
            "QPushButton { background: #059669; color: #FFF; border: none; border-radius: 8px; padding: 0 20px; }"
            "QPushButton:hover { background: #047857; }"
        )
        btn_save.clicked.connect(self._on_save)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_save)
        lay.addLayout(btn_row)

    def _on_preview(self) -> None:
        if not self._db or not self._db.is_connected() or not self._student_id:
            show_warning(self, "미리보기", "DB 연결 또는 학생 정보가 없습니다.")
            return
        start_d = self._date_start.date()
        end_d = self._date_end.date()
        if start_d > end_d:
            show_warning(self, "미리보기", "시작일이 종료일보다 늦을 수 없습니다.")
            return
        start_s = start_d.toString("yyyy-MM-dd")
        end_s = end_d.toString("yyyy-MM-dd")
        comment = (self._comment_input.toPlainText() or "").strip()
        try:
            snapshot = aggregate_report(self._db, self._student_id, start_s, end_s)
        except Exception as e:
            show_warning(self, "미리보기 오류", f"보고서 집계 중 오류가 발생했습니다.\n\n{type(e).__name__}: {e}")
            return
        self._current_snapshot = snapshot
        total_q = int(snapshot.get("total_questions") or 0)
        if total_q == 0:
            show_info(self, "미리보기", "해당 기간에 채점된 학습지가 없습니다.")
        try:
            dashboard = _build_report_dashboard(
                snapshot,
                self._student_name,
                start_s,
                end_s,
                comment=comment,
            )
        except Exception as e:
            show_warning(self, "미리보기 오류", f"보고서 화면 생성 중 오류가 발생했습니다.\n\n{type(e).__name__}: {e}")
            return
        old = self._preview_scroll.widget()
        if old is not None:
            try:
                old.setParent(None)
            except Exception:
                pass
        self._preview_scroll.setWidget(dashboard)
        # setWidget 후 이전 위젯은 Qt에 의해 삭제될 수 있으므로
        # _preview_placeholder 등 이전 위젯 참조 금지 (RuntimeError: wrapped C/C++ object has been deleted 방지)

    def _on_save(self) -> None:
        if not self._current_snapshot or not self._db or not self._db.is_connected() or not self._student_id:
            show_warning(self, "저장", "먼저 '미리보기'를 눌러 보고서를 불러온 뒤 저장해 주세요.")
            return
        start_s = self._date_start.date().toString("yyyy-MM-dd")
        end_s = self._date_end.date().toString("yyyy-MM-dd")
        comment = (self._comment_input.toPlainText() or "").strip()
        report = SavedReport(
            id=None,
            student_id=self._student_id,
            period_start=start_s,
            period_end=end_s,
            comment=comment,
            created_at=None,
            snapshot=dict(self._current_snapshot),
        )
        try:
            repo = ReportRepository(self._db)
            rid = repo.create(report)
        except Exception as e:
            show_warning(self, "저장 실패", str(e))
            return
        show_info(self, "저장 완료", "보고서가 저장되었습니다.")
        self.accept()


class ReportCommentEditDialog(QDialog):
    """보고서 수정 시 학습코멘트만 편집하는 다이얼로그. QTextEdit로 넉넉한 입력 공간 제공."""

    def __init__(self, parent: Optional[QWidget], current_comment: str) -> None:
        super().__init__(parent)
        self.setWindowTitle("학습코멘트 수정")
        self.setMinimumWidth(480)
        self.resize(520, 320)
        self.setStyleSheet("QDialog { background-color: #F8FAFC; }")

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)

        lay.addWidget(QLabel("학습코멘트:"))
        self._comment_input = QTextEdit()
        self._comment_input.setPlaceholderText("학습 코멘트를 입력하세요. 여러 줄 입력 가능합니다.")
        self._comment_input.setPlainText((current_comment or "").strip())
        self._comment_input.setMinimumHeight(120)
        self._comment_input.setMaximumHeight(200)
        self._comment_input.setStyleSheet(
            "QTextEdit { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 10px; font-size: 11pt; }"
        )
        lay.addWidget(self._comment_input)

        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        btn_cancel = QPushButton("취소")
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setFixedHeight(36)
        btn_cancel.setStyleSheet(
            "QPushButton { background: #E2E8F0; color: #334155; border: none; border-radius: 8px; padding: 0 20px; }"
            "QPushButton:hover { background: #CBD5E1; }"
        )
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("확인")
        btn_ok.setCursor(Qt.PointingHandCursor)
        btn_ok.setFixedHeight(36)
        btn_ok.setStyleSheet(
            "QPushButton { background: #059669; color: #FFF; border: none; border-radius: 8px; padding: 0 20px; }"
            "QPushButton:hover { background: #047857; }"
        )
        btn_ok.clicked.connect(self._on_accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        lay.addLayout(btn_row)

    def _on_accept(self) -> None:
        self.accept()

    def get_comment(self) -> str:
        return self._comment_input.toPlainText().strip()


class ReportDetailModal(QDialog):
    """보고서 리스트에서 행 클릭 시 열리는 모달: 상세 내용 + 하단 PDF / 닫기 버튼."""

    def __init__(
        self,
        report: SavedReport,
        student_name: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.report = report
        self.student_name = (student_name or "").strip()
        self.setWindowTitle("학습 보고서")
        self.setMinimumSize(1000, 700)
        self.resize(1100, 800)
        self.setStyleSheet("QDialog { background-color: #FFFFFF; }")
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        snapshot = dict(report.snapshot or {})
        created_str = ""
        if report.created_at:
            try:
                created_str = report.created_at.strftime("%Y.%m.%d %H:%M")
            except Exception:
                created_str = str(report.created_at)
        dashboard = _build_report_dashboard(
            snapshot,
            self.student_name,
            report.period_start or "",
            report.period_end or "",
            created_at=created_str,
            comment=report.comment or "",
        )
        self._scroll_area = dashboard
        lay.addWidget(dashboard, 1)
        btn_bar = QFrame()
        btn_bar.setStyleSheet(
            "QFrame { background: #FFFFFF; border-top: 1px solid #EAEAEA; padding: 12px; }"
        )
        btn_lay = QHBoxLayout(btn_bar)
        btn_lay.setContentsMargins(24, 12, 24, 12)
        btn_lay.setSpacing(12)
        btn_lay.addStretch(1)
        btn_pdf = QPushButton("PDF")
        btn_pdf.setCursor(Qt.PointingHandCursor)
        btn_pdf.setFixedHeight(40)
        btn_pdf.setMinimumWidth(120)
        btn_pdf.setStyleSheet(
            """
            QPushButton {
                background: #1976D2;
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                font-size: 12pt;
                font-weight: bold;
            }
            QPushButton:hover { background: #1565C0; }
            QPushButton:pressed { background: #0D47A1; }
            """
        )
        btn_pdf.clicked.connect(self._on_pdf)
        btn_close = QPushButton("닫기")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedHeight(40)
        btn_close.setMinimumWidth(100)
        btn_close.setStyleSheet(
            """
            QPushButton {
                background: #F5F5F5;
                color: #222222;
                border: 1px solid #E0E0E0;
                border-radius: 8px;
                font-size: 12pt;
            }
            QPushButton:hover { background: #EEEEEE; }
            """
        )
        btn_close.clicked.connect(self.accept)
        btn_lay.addWidget(btn_pdf)
        btn_lay.addWidget(btn_close)
        lay.addWidget(btn_bar)

    def _on_pdf(self) -> None:
        """PDF 버튼 클릭 시 먼저 '저장/열기/취소' 모달을 띄운 뒤, 선택에 따라 저장 또는 열기."""
        if not _HAS_QPRINTER:
            show_warning(self, "PDF", "PDF 출력을 위해 PyQt5.QtPrintSupport가 필요합니다.")
            return
        choice = StandardActionDialog(
            parent=self,
            title="PDF",
            message="PDF 작업을 선택하세요.",
            actions=[
                DialogAction(key="save", label="저장", is_primary=False),
                DialogAction(key="open", label="열기", is_primary=True),
                DialogAction(key="cancel", label="취소", is_primary=False),
            ],
            min_width=380,
        )
        choice.exec_()
        if choice.selected_key not in ("save", "open"):
            return

        scroll = self._scroll_area
        if not scroll or not scroll.widget():
            show_warning(self, "PDF", "보고서 내용을 찾을 수 없습니다.")
            return
        wrapper = scroll.widget()
        wlay = wrapper.layout()
        content_w = 1000
        # 1000px 레이아웃을 2.4배 확대 → 2400px (기존 1200×2.0과 동일), PDF에서 요소 크기 유지
        capture_zoom = 2.4

        container = None
        container_old_max_w = None
        if wlay and wlay.count() >= 2:
            cont_item = wlay.itemAt(1)
            if cont_item and cont_item.widget():
                container = cont_item.widget()
                container_old_max_w = container.maximumWidth()
                container.setMaximumWidth(content_w)
        QApplication.processEvents()
        QApplication.processEvents()

        # wrapper를 크게 잡아 레이아웃이 전체 높이로 계산되게 한 뒤, container 기준으로 content_h 확보
        wrapper.resize(content_w, 50000)
        QApplication.processEvents()
        QApplication.processEvents()

        vb = scroll.verticalScrollBar()
        full_h = vb.maximum() + scroll.viewport().height()
        content_h = max(600, full_h)
        if container is not None:
            try:
                h = container.minimumSizeHint().height()
                if h > 0:
                    content_h = max(content_h, h)
            except Exception:
                pass
            try:
                h = container.layout().minimumSize().height()
                if h > 0:
                    content_h = max(content_h, h)
            except Exception:
                pass
        content_h = max(600, content_h)

        old_size = wrapper.size()
        wrapper.resize(content_w, content_h)
        QApplication.processEvents()
        QApplication.processEvents()
        try:
            # 1000px 레이아웃을 확대 픽셀맵에 그려 선명한 캡처 (요소 크기 유지)
            pix_w_zoomed = int(content_w * capture_zoom)
            pix_h_zoomed = int(content_h * capture_zoom)
            pixmap = QPixmap(pix_w_zoomed, pix_h_zoomed)
            pixmap.fill(Qt.white)
            painter = QPainter(pixmap)
            painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
            painter.scale(capture_zoom, capture_zoom)
            wrapper.render(painter)
            painter.end()
        finally:
            wrapper.resize(old_size)
            if container_old_max_w is not None and wlay and wlay.count() >= 2:
                cont_item = wlay.itemAt(1)
                if cont_item and cont_item.widget():
                    cont_item.widget().setMaximumWidth(container_old_max_w)
        if pixmap.isNull():
            show_warning(self, "PDF", "보고서 캡처에 실패했습니다.")
            return

        period_s = (self.report.period_start or "").replace("-", ".")
        period_e = (self.report.period_end or "").replace("-", ".")
        default_name = f"학습보고서_{self.student_name}_{period_s}_{period_e}.pdf".replace(" ", "_")

        if choice.selected_key == "save":
            default_path = os.path.join(os.path.expanduser("~"), "Desktop", default_name)
            path, _ = QFileDialog.getSaveFileName(
                self, "PDF 저장", default_path, "PDF Files (*.pdf)"
            )
            if not path or not path.strip():
                return
            path = path.strip()
            if not path.lower().endswith(".pdf"):
                path += ".pdf"
        else:
            tmp_dir = _ascii_temp_dir("report_pdf")
            path = os.path.join(tmp_dir, default_name)

        try:
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(path)
            printer.setPageMargins(8, 8, 8, 8, QPrinter.Millimeter)
            page_rect = printer.pageRect(QPrinter.DevicePixel)
            page_w = page_rect.width()
            page_h = page_rect.height()
            bottom_margin_px = 48
            usable_page_h = max(100, page_h - bottom_margin_px)

            pix_w = pixmap.width()
            pix_h = pixmap.height()
            if pix_w <= 0 or pix_h <= 0:
                show_warning(self, "PDF", "캡처 크기가 올바르지 않습니다.")
                return
            scale = page_w / float(pix_w)
            scaled_h = pix_h * scale
            num_pages = max(1, int((scaled_h / usable_page_h) + 0.99))
            painter = QPainter(printer)
            painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
            for i in range(num_pages):
                if i > 0:
                    printer.newPage()
                src_y = int(i * usable_page_h / scale)
                src_h = min(int(usable_page_h / scale), pix_h - src_y)
                if src_h <= 0:
                    break
                src_rect = QRect(0, src_y, pix_w, src_h)
                dest_h = min(usable_page_h, src_h * scale)
                dest_rect = QRectF(0, 0, page_w, dest_h)
                painter.drawPixmap(dest_rect, pixmap, QRectF(src_rect))
            painter.end()
        except Exception as e:
            show_warning(self, "PDF 저장 실패", str(e))
            return

        if choice.selected_key == "open":
            try:
                os.startfile(path)  # type: ignore[attr-defined]
            except Exception:
                show_warning(self, "열기 실패", "PDF를 기본 프로그램으로 열 수 없습니다.")
        else:
            dlg = StandardActionDialog(
                parent=self,
                title="PDF 저장",
                message="PDF 파일이 저장되었습니다. 파일을 열까요?",
                actions=[
                    DialogAction(key="close", label="저장만", is_primary=False),
                    DialogAction(key="open", label="열기", is_primary=True),
                ],
                min_width=400,
            )
            dlg.exec_()
            if dlg.selected_key == "open":
                try:
                    os.startfile(path)  # type: ignore[attr-defined]
                except Exception:
                    show_warning(self, "열기 실패", "PDF를 기본 프로그램으로 열 수 없습니다.")


# ==============================================================================
# [프리미엄] 보고서 상세: Full-Page Scroll + 중앙 고정 너비 + 카드 그림자/radius 20px
# 리스트 화면은 건드리지 않고 상세화면만 적용.
# ==============================================================================

_REPORT_HIGH = "#D32F2F"
_REPORT_MID = "#1976D2"
_REPORT_LOW = "#388E3C"

_REPORT_CARD_STYLE = """
    QFrame {
        background-color: #FFFFFF;
        border: 1px solid #EAEAEA;
        border-radius: 20px;
    }
"""
_REPORT_SHADOW_BLUR = 25
_REPORT_SHADOW_OFFSET = 8
_REPORT_SHADOW_ALPHA = 20  # rgba(0,0,0,0.08) ≈ 20


def _format_report_period(start: str, end: str) -> str:
    def _fmt(s: str) -> str:
        if not s or len(s) < 10:
            return s
        try:
            y, m, d = s[:4], s[5:7], s[8:10]
            return f"{y}년 {int(m)}월 {int(d)}일"
        except Exception:
            return s.replace("-", ".")
    return f"{_fmt(start)} ~ {_fmt(end)}"


def _add_report_card_shadow(widget: QWidget) -> None:
    shadow = QGraphicsDropShadowEffect(widget)
    shadow.setBlurRadius(_REPORT_SHADOW_BLUR)
    shadow.setOffset(0, _REPORT_SHADOW_OFFSET)
    shadow.setColor(QColor(0, 0, 0, _REPORT_SHADOW_ALPHA))
    widget.setGraphicsEffect(shadow)


# --- 완료한 과제 도넛 (종합 학습 분석용) ---
_DONUT_PINK = QColor(244, 143, 177)   # #F48FB1
_DONUT_PURPLE = QColor(156, 39, 176)  # #9C27B0
_DONUT_LIGHT_PURPLE = QColor(206, 147, 216)  # #CE93D8


class CompletedAssignmentsDonut(QFrame):
    """완료한 과제 수를 도넛 링 + 중앙 텍스트로 표시 (그라데이션 링)."""
    def __init__(self, count: int, parent: Optional[QWidget] = None, *, embedded: bool = False):
        super().__init__(parent)
        self._count = max(0, int(count))
        if embedded:
            self.setStyleSheet("background: transparent; border: none;")
        else:
            self.setStyleSheet(_REPORT_CARD_STYLE)
            _add_report_card_shadow(self)
        self.setMinimumSize(200, 200)
        self.setMaximumWidth(280)

        text_container = QWidget(self)
        text_container.setStyleSheet("background: transparent; border: none;")
        text_layout = QVBoxLayout(text_container)
        text_layout.setSpacing(2)
        text_layout.setContentsMargins(0, 0, 0, 0)
        self._count_label = QLabel(f"{self._count}개")
        self._count_label.setStyleSheet("font-size: 18pt; font-weight: bold; color: #000000; border: none; background: transparent;")
        self._count_label.setAlignment(Qt.AlignCenter)
        title_label = QLabel("완료한 과제")
        title_label.setStyleSheet("font-size: 9pt; color: #555555; border: none; background: transparent;")
        title_label.setAlignment(Qt.AlignCenter)
        text_layout.addWidget(self._count_label)
        text_layout.addWidget(title_label)

        center_layout = QVBoxLayout(self)
        center_layout.setContentsMargins(0, 0, 0, 0)
        center_layout.setSpacing(0)
        center_layout.addStretch(1)
        center_layout.addWidget(text_container, 0, Qt.AlignCenter)
        center_layout.addStretch(1)

    def set_count(self, count: int) -> None:
        self._count = max(0, int(count))
        self._count_label.setText(f"{self._count}개")
        self.update()

    def paintEvent(self, event) -> None:
        super().paintEvent(event)
        w, h = self.width(), self.height()
        if w <= 0 or h <= 0:
            return
        side = min(w, h)
        margin = 20
        cx, cy = w / 2, h / 2
        outer_r = (side / 2) - margin
        inner_r = outer_r * 0.55
        if outer_r <= inner_r or outer_r <= 0:
            return
        path = QPainterPath()
        path.addEllipse(cx - outer_r, cy - outer_r, outer_r * 2, outer_r * 2)
        path.addEllipse(cx - inner_r, cy - inner_r, inner_r * 2, inner_r * 2)
        path.setFillRule(Qt.OddEvenFill)
        gradient = QConicalGradient(cx, cy, 90)
        gradient.setColorAt(0.0, _DONUT_PINK)
        gradient.setColorAt(0.5, _DONUT_PURPLE)
        gradient.setColorAt(0.75, _DONUT_LIGHT_PURPLE)
        gradient.setColorAt(1.0, _DONUT_PINK)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.setRenderHint(QPainter.SmoothPixmapTransform, True)
        painter.fillPath(path, QBrush(gradient))
        painter.end()


# --- 단원별 성취도 카드 (소단원 분석용) ---
class NewReportCardWidget(QFrame):
    def __init__(
        self,
        unit_name: str,
        total_score: float,
        details: Dict[str, float],
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.setStyleSheet(_REPORT_CARD_STYLE + " QFrame QLabel { color: #222222; border: none; background: transparent; }")
        _add_report_card_shadow(self)
        self.setFixedHeight(168)
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(8)

        header_layout = QHBoxLayout()
        unit_label = QLabel(unit_name)
        unit_label.setStyleSheet("font-size: 13pt; font-weight: bold; color: #000000; border: none; background: transparent;")
        score_label = QLabel(f"평균 {total_score:.1f}%")
        score_label.setStyleSheet("font-size: 11pt; color: #222222; border: none; background: transparent;")
        header_layout.addWidget(unit_label)
        header_layout.addStretch(1)
        header_layout.addWidget(score_label)
        main_layout.addLayout(header_layout)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("background: #EAEAEA; max-height: 1px; border: none;")
        main_layout.addWidget(line)

        color_map = {
            "상": (_REPORT_GRADIENT_HIGH, _REPORT_HIGH, "High"),
            "중": (_REPORT_GRADIENT_MID, _REPORT_MID, "Mid"),
            "하": (_REPORT_GRADIENT_LOW, _REPORT_LOW, "Low"),
        }
        for level in ["상", "중", "하"]:
            score = max(0, min(100, float(details.get(level, 0))))
            gradient_style, color_code, eng = color_map[level]
            row = QHBoxLayout()
            lbl = QLabel(f"{level} ({eng})")
            lbl.setFixedWidth(68)
            lbl.setStyleSheet("font-weight: bold; font-size: 10pt; color: #222222; border: none; background: transparent;")
            pbar = QProgressBar()
            pbar.setMaximum(100)
            pbar.setValue(int(round(score)))
            pbar.setFixedHeight(10)
            pbar.setTextVisible(False)
            pbar.setStyleSheet(
                "QProgressBar { background-color: #F5F5F5; border-radius: 5px; border: none; }"
                f"QProgressBar::chunk {{ background: {gradient_style}; border-radius: 5px; }}"
            )
            score_lbl = QLabel(f"{int(round(score))}%")
            score_lbl.setFixedWidth(40)
            score_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            score_lbl.setStyleSheet(f"color: {color_code}; font-weight: bold; font-size: 10pt; border: none; background: transparent;")
            row.addWidget(lbl)
            row.addWidget(pbar, 1)
            row.addWidget(score_lbl)
            main_layout.addLayout(row)


# --- 교재별 분석 (교재 출처별 N개 중 M개 맞음) ---
class TextbookItemWidget(QWidget):
    """교재별 분석 한 행: 교재명 + N개 중 M개 맞음."""
    def __init__(self, name: str, correct: int, total: int, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent; border: none;")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 15, 18, 15)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignVCenter)

        name_lbl = QLabel(name or "미명")
        name_lbl.setStyleSheet(
            "font-size: 12pt; font-weight: 500; color: #000000; border: none; background: transparent;"
        )
        name_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        score_lbl = QLabel(f"{total}개 중 <b>{correct}개</b> 맞음")
        score_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        score_lbl.setTextFormat(Qt.RichText)
        score_lbl.setStyleSheet("font-size: 12pt; color: #333333; border: none; background: transparent;")

        layout.addWidget(name_lbl, 1)
        layout.addWidget(score_lbl)


def _report_textbook_section(snapshot: dict, parent: Optional[QWidget] = None) -> QWidget:
    """교재별 분석 섹션: 교재 출처별로 N개 중 M개 맞음 표시."""
    textbook_stats = list(snapshot.get("textbook_stats") or [])
    if not textbook_stats and snapshot.get("source_stats"):
        for item in snapshot.get("source_stats") or []:
            if str(item.get("category") or "").strip() != "기출":
                textbook_stats.append({
                    "name": item.get("name"),
                    "correct": item.get("correct"),
                    "total": item.get("total"),
                })

    card = QFrame(parent)
    card.setStyleSheet(_REPORT_CARD_STYLE + " QLabel { border: none; background: transparent; }")
    _add_report_card_shadow(card)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(24, 20, 24, 24)
    lay.setSpacing(0)

    title = QLabel("📚 교재별 분석")
    title.setStyleSheet("font-size: 18pt; font-weight: bold; color: #000000; margin-bottom: 12px; border: none; background: transparent;")
    lay.addWidget(title)

    if not textbook_stats:
        empty = QLabel("교재별 데이터가 없습니다.")
        empty.setStyleSheet("font-size: 11pt; color: #666666; padding: 20px 0; border: none; background: transparent;")
        empty.setAlignment(Qt.AlignCenter)
        lay.addWidget(empty)
        return card

    for i, item in enumerate(textbook_stats):
        if i > 0:
            line = QFrame()
            line.setFrameShape(QFrame.HLine)
            line.setStyleSheet("background: #F5F5F5; max-height: 1px; border: none; margin: 0 0 0 0;")
            lay.addWidget(line)
        name = str(item.get("name") or "미명").strip()
        correct = int(item.get("correct") or 0)
        total = int(item.get("total") or 0)
        lay.addWidget(TextbookItemWidget(name, correct, total))

    return card


# --- 기출별 분석 (기출별 M개 맞음, K개 틀림) ---
class ExamItemWidget(QWidget):
    """기출별 분석 한 행: 기출명(예: OO학교 O학년 O학기 중간고사 2026) + M개 맞음, K개 틀림."""
    def __init__(self, name: str, correct: int, total: int, parent: Optional[QWidget] = None):
        super().__init__(parent)
        wrong = max(0, total - correct)
        self.setStyleSheet("background: transparent; border: none;")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(18, 15, 18, 15)
        layout.setSpacing(10)
        layout.setAlignment(Qt.AlignVCenter)

        name_lbl = QLabel(name or "미명")
        name_lbl.setStyleSheet(
            "font-size: 12pt; font-weight: 500; color: #000000; border: none; background: transparent;"
        )
        name_lbl.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        score_lbl = QLabel(f"<b>{correct}개</b> 맞음, <b>{wrong}개</b> 틀림")
        score_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        score_lbl.setTextFormat(Qt.RichText)
        score_lbl.setStyleSheet("font-size: 12pt; color: #333333; border: none; background: transparent;")

        layout.addWidget(name_lbl, 1)
        layout.addWidget(score_lbl)


def _report_exam_section(snapshot: dict, parent: Optional[QWidget] = None) -> QWidget:
    """기출별 분석 섹션: 기출(예: OO학교 O학년 O학기 중간고사 2026)별 맞음/틀림 표시."""
    exam_stats = list(snapshot.get("exam_stats") or [])
    if not exam_stats and snapshot.get("source_stats"):
        for item in snapshot.get("source_stats") or []:
            if str(item.get("category") or "").strip() == "기출":
                exam_stats.append({
                    "name": item.get("name"),
                    "correct": item.get("correct"),
                    "total": item.get("total"),
                })

    card = QFrame(parent)
    card.setStyleSheet(_REPORT_CARD_STYLE + " QLabel { border: none; background: transparent; }")
    _add_report_card_shadow(card)
    lay = QVBoxLayout(card)
    lay.setContentsMargins(24, 20, 24, 24)
    lay.setSpacing(0)

    title = QLabel("📋 기출별 분석")
    title.setStyleSheet("font-size: 18pt; font-weight: bold; color: #000000; margin-bottom: 12px; border: none; background: transparent;")
    lay.addWidget(title)

    if not exam_stats:
        empty = QLabel("기출별 데이터가 없습니다.")
        empty.setStyleSheet("font-size: 11pt; color: #666666; padding: 20px 0; border: none; background: transparent;")
        empty.setAlignment(Qt.AlignCenter)
        lay.addWidget(empty)
        return card

    for i, item in enumerate(exam_stats):
        if i > 0:
            line = QFrame()
            line.setFrameShape(QFrame.HLine)
            line.setStyleSheet("background: #F5F5F5; max-height: 1px; border: none; margin: 0 0 0 0;")
            lay.addWidget(line)
        name = str(item.get("name") or "미명").strip()
        correct = int(item.get("correct") or 0)
        total = int(item.get("total") or 0)
        lay.addWidget(ExamItemWidget(name, correct, total))

    return card


# --- 종합 학습 분석 (레이더) ---
def _report_radar_widget(unit_stats: List[Dict], parent: Optional[QWidget] = None) -> QWidget:
    wrap = QFrame(parent)
    wrap.setStyleSheet(_REPORT_CARD_STYLE + " QLabel { border: none; background: transparent; }")
    _add_report_card_shadow(wrap)
    wrap.setMinimumHeight(300)
    lay = QVBoxLayout(wrap)
    lay.setContentsMargins(30, 30, 30, 30)
    if not _HAS_MATPLOTLIB or not unit_stats:
        lbl = QLabel("종합 학습 분석\n(데이터 없음)" if not unit_stats else "종합 학습 분석")
        lbl.setAlignment(Qt.AlignCenter)
        lbl.setStyleSheet("font-size: 12pt; color: #222222; border: none; background: transparent;")
        lay.addWidget(lbl)
        return wrap
    import math
    stats = unit_stats[:6]
    labels = [str(s.get("unit_label") or s.get("unit_key") or "—")[:8] for s in stats]
    values = [float(s.get("rate_pct") or 0) for s in stats]
    while len(labels) < 3:
        labels.append("—")
        values.append(0)
    n = len(labels)
    angles = [2 * math.pi * i / n - math.pi / 2 for i in range(n)]
    values_c = values + values[:1]
    angles_c = angles + angles[:1]
    fig = Figure(figsize=(2.5, 2.5), facecolor="none")
    ax = fig.add_subplot(111, polar=True)
    ax.set_facecolor("none")
    ax.patch.set_visible(False)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.grid(True, color="#F0F0F0", linewidth=0.8, alpha=0.9)
    ax.plot(angles_c, values_c, "o-", linewidth=2, color=_REPORT_MID, markersize=8, markerfacecolor=_REPORT_MID, markeredgecolor="#fff", markeredgewidth=1.2)
    ax.fill(angles_c, values_c, alpha=0.28, color=_REPORT_MID)
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, size=8)
    ax.set_ylim(0, 100)
    ax.tick_params(colors="#222222")
    fig.tight_layout(pad=0.5)
    canvas = FigureCanvasQTAgg(fig)
    canvas.setStyleSheet("background: transparent;")
    lay.addWidget(canvas)
    return wrap


# --- 학습 난이도 분석 (상/중/하 막대, 그라데이션) ---
_REPORT_GRADIENT_HIGH = "qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #B71C1C, stop:1 #FFCDD2)"
_REPORT_GRADIENT_MID = "qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #0D47A1, stop:1 #BBDEFB)"
_REPORT_GRADIENT_LOW = "qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #1B5E20, stop:1 #C8E6C9)"


def _report_difficulty_content(
    rate_high: float,
    rate_mid: float,
    rate_low: float,
    parent: Optional[QWidget] = None,
) -> QWidget:
    """난이도 분석 내용만 반환(카드 없음, 통합 분석 카드 내부용)."""
    content = QWidget(parent)
    content.setStyleSheet("QLabel { border: none; background: transparent; }")
    lay = QVBoxLayout(content)
    lay.setContentsMargins(24, 24, 24, 24)
    lay.setSpacing(15)
    lay.setAlignment(Qt.AlignVCenter)

    lay.addStretch(1)
    title = QLabel("학습 난이도 분석")
    title.setStyleSheet("font-size: 14pt; font-weight: bold; color: #000000; margin-bottom: 8px;")
    lay.addWidget(title)
    desc = QLabel("난이도별 정답률을 통해 현재 어떤 수준의 문제를 해결할 수 있는지 진단합니다.")
    desc.setWordWrap(True)
    desc.setStyleSheet("font-size: 11pt; color: #222222; margin-bottom: 16px;")
    lay.addWidget(desc)
    levels = [
        ("상 (High)", max(0, min(100, rate_high)), _REPORT_GRADIENT_HIGH, _REPORT_HIGH),
        ("중 (Mid)", max(0, min(100, rate_mid)), _REPORT_GRADIENT_MID, _REPORT_MID),
        ("하 (Low)", max(0, min(100, rate_low)), _REPORT_GRADIENT_LOW, _REPORT_LOW),
    ]
    for label_text, score, gradient_style, color_code in levels:
        row = QHBoxLayout()
        row.setAlignment(Qt.AlignVCenter)
        lbl = QLabel(label_text)
        lbl.setFixedWidth(90)
        lbl.setStyleSheet("font-weight: bold; font-size: 11pt; color: #222222; border: none; background: transparent;")
        pbar = QProgressBar()
        pbar.setMaximum(100)
        pbar.setValue(int(round(score)))
        pbar.setFixedHeight(14)
        pbar.setTextVisible(False)
        pbar.setStyleSheet(
            "QProgressBar { background-color: #F5F5F5; border-radius: 7px; border: none; }"
            f"QProgressBar::chunk {{ background: {gradient_style}; border-radius: 7px; }}"
        )
        pct = QLabel(f"{int(round(score))}%")
        pct.setFixedWidth(45)
        pct.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        pct.setStyleSheet(f"color: {color_code}; font-weight: bold; font-size: 11pt; border: none; background: transparent;")
        row.addWidget(lbl)
        row.addWidget(pbar, 1)
        row.addWidget(pct)
        lay.addLayout(row)
    lay.addStretch(1)
    return content


# --- 메인: Full-Page Scroll + 중앙 고정(최대 1200px) + 모든 섹션 한 스크롤 ---
def _build_report_dashboard(
    snapshot: dict,
    student_name: str,
    period_start: str,
    period_end: str,
    *,
    created_at: Optional[str] = None,
    comment: Optional[str] = None,
    student_grade: Optional[str] = None,
    school_name: Optional[str] = None,
    parent: Optional[QWidget] = None,
) -> QWidget:
    scroll_area = QScrollArea(parent)
    scroll_area.setWidgetResizable(True)
    scroll_area.setFrameShape(QFrame.NoFrame)
    scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #FFFFFF; }")
    scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    wrapper = QWidget()
    wrapper.setStyleSheet("background-color: #FFFFFF;")
    wrapper.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
    wrapper_lay = QHBoxLayout(wrapper)
    wrapper_lay.setContentsMargins(0, 0, 0, 0)
    wrapper_lay.setSpacing(0)

    container = QWidget()
    container.setStyleSheet(
        "background-color: #FFFFFF;"
        " QLabel { border: none; background: transparent; }"
        " QFrame#AnalysisCard { background-color: white; border-radius: 20px; border: 1px solid #EAEAEA; }"
    )
    container.setMaximumWidth(1200)
    container.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Minimum)
    main_layout = QVBoxLayout(container)
    main_layout.setContentsMargins(60, 40, 60, 40)
    main_layout.setSpacing(24)

    header_container = QWidget()
    header_container.setStyleSheet("background: transparent; border: none;")
    header_layout = QVBoxLayout(header_container)
    header_layout.setAlignment(Qt.AlignCenter)
    header_layout.setSpacing(5)
    header_layout.setContentsMargins(0, 0, 0, 0)

    student_display = (student_name or "학생").strip()
    name_label = QLabel(f"{student_display}학생의")
    name_label.setStyleSheet(
        "font-size: 28pt; font-weight: 900; color: #000000;"
        " border: none; background: transparent;"
    )
    name_label.setAlignment(Qt.AlignCenter)

    period_str = _format_report_period(period_start or "", period_end or "")
    period_label = QLabel(f"{period_str} 학습 보고서")
    period_label.setStyleSheet(
        "font-size: 14pt; font-weight: 500; color: #000000;"
        " border: none; background: transparent;"
    )
    period_label.setAlignment(Qt.AlignCenter)

    header_layout.addWidget(name_label)
    header_layout.addWidget(period_label)
    main_layout.addWidget(header_container)

    comment_card = QFrame()
    comment_card.setStyleSheet(
        _REPORT_CARD_STYLE + " QFrame QLabel { border: none; background: transparent; color: #222222; }"
    )
    _add_report_card_shadow(comment_card)
    c_layout = QVBoxLayout(comment_card)
    c_layout.setContentsMargins(30, 30, 30, 30)
    c_layout.setSpacing(12)
    c_title = QLabel("학습 Comment")
    c_title.setStyleSheet("font-weight: bold; font-size: 14pt; color: #000000; border: none; background: transparent;")
    c_content = QLabel(comment.strip() if comment else "학습 코멘트가 없습니다.")
    c_content.setWordWrap(True)
    c_content.setStyleSheet("font-size: 12pt; color: #222222; line-height: 1.6; border: none; background: transparent;")
    c_layout.addWidget(c_title)
    c_layout.addWidget(c_content)
    main_layout.addWidget(comment_card)

    unit_stats = list(snapshot.get("unit_stats") or [])
    total_worksheets = int(snapshot.get("total_worksheets") or 0)
    rate = float(snapshot.get("average_rate_pct") or 0)

    analysis_card = QFrame()
    analysis_card.setObjectName("AnalysisCard")
    analysis_card.setStyleSheet(_REPORT_CARD_STYLE + " QLabel { border: none; background: transparent; }")
    _add_report_card_shadow(analysis_card)
    analysis_card.setMinimumHeight(280)
    analysis_card_lay = QHBoxLayout(analysis_card)
    analysis_card_lay.setContentsMargins(16, 16, 16, 16)
    analysis_card_lay.setSpacing(24)
    analysis_card_lay.setAlignment(Qt.AlignVCenter)
    donut_wrap = QWidget()
    donut_wrap.setStyleSheet("background: transparent; border: none;")
    donut_wrap.setFixedHeight(260)
    donut_wrap_lay = QVBoxLayout(donut_wrap)
    donut_wrap_lay.setContentsMargins(0, 0, 0, 0)
    donut_wrap_lay.setAlignment(Qt.AlignCenter)
    donut_wrap_lay.addWidget(CompletedAssignmentsDonut(total_worksheets, embedded=True))
    analysis_card_lay.addWidget(donut_wrap, 0, Qt.AlignVCenter)
    difficulty_content = _report_difficulty_content(rate * 0.88, rate, min(100, rate * 1.12))
    difficulty_content.setMinimumHeight(260)
    analysis_card_lay.addWidget(difficulty_content, 1, Qt.AlignVCenter)
    main_layout.addWidget(analysis_card, 0)

    unit_section_label = QLabel("소단원 분석")
    unit_section_label.setStyleSheet("font-size: 14pt; font-weight: bold; color: #000000; border: none; background: transparent;")
    main_layout.addWidget(unit_section_label)

    for u in unit_stats:
        rate_pct = float(u.get("rate_pct") or 0)
        details = {
            "상": float(u.get("rate_high") or rate_pct * 0.88),
            "중": float(u.get("rate_mid") or rate_pct),
            "하": min(100.0, float(u.get("rate_low") or rate_pct * 1.12)),
        }
        unit_name = str(u.get("unit_label") or u.get("unit_key") or "미분류")
        main_layout.addWidget(NewReportCardWidget(unit_name, rate_pct, details))

    main_layout.addWidget(_report_textbook_section(snapshot))
    main_layout.addWidget(_report_exam_section(snapshot))

    wrapper_lay.addStretch(1)
    wrapper_lay.addWidget(container)
    wrapper_lay.addStretch(1)
    scroll_area.setWidget(wrapper)
    return scroll_area


def _build_snapshot_widget(snapshot: dict, parent: Optional[QWidget] = None) -> QWidget:
    """집계 스냅샷을 표시하는 위젯(요약 + 단원별 테이블)."""
    wrap = QWidget(parent)
    lay = QVBoxLayout(wrap)
    lay.setSpacing(12)

    total_ws = int(snapshot.get("total_worksheets") or 0)
    total_q = int(snapshot.get("total_questions") or 0)
    total_c = int(snapshot.get("total_correct") or 0)
    rate = float(snapshot.get("average_rate_pct") or 0)

    summary = QFrame()
    summary.setObjectName("ReportSummaryCard")
    summary.setStyleSheet(
        """
        QFrame#ReportSummaryCard {
            background: #EFF6FF;
            border: 1px solid #BFDBFE;
            border-radius: 8px;
            padding: 12px;
        }
        """
    )
    slay = QGridLayout(summary)
    slay.addWidget(QLabel("총 학습지 수"), 0, 0)
    slay.addWidget(QLabel(str(total_ws)), 0, 1)
    slay.addWidget(QLabel("총 문항 수"), 1, 0)
    slay.addWidget(QLabel(str(total_q)), 1, 1)
    slay.addWidget(QLabel("총 정답 수"), 2, 0)
    slay.addWidget(QLabel(str(total_c)), 2, 1)
    slay.addWidget(QLabel("평균 정답률"), 3, 0)
    slay.addWidget(QLabel(f"{rate:.1f}%"), 3, 1)
    lay.addWidget(summary)

    unit_stats = list(snapshot.get("unit_stats") or [])
    if unit_stats:
        ulabel = QLabel("단원별 정답률")
        ulabel.setFont(_font(10, bold=True))
        ulabel.setStyleSheet("color: #0F172A;")
        lay.addWidget(ulabel)
        table_wrap = QFrame()
        table_wrap.setStyleSheet(
            """
            QFrame { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 6px; }
            """
        )
        tlay = QGridLayout(table_wrap)
        tlay.addWidget(QLabel("단원"), 0, 0)
        tlay.addWidget(QLabel("총 문항"), 0, 1)
        tlay.addWidget(QLabel("정답"), 0, 2)
        tlay.addWidget(QLabel("정답률"), 0, 3)
        for i, u in enumerate(unit_stats, start=1):
            tlay.addWidget(QLabel(str(u.get("unit_label") or u.get("unit_key") or "—")), i, 0)
            tlay.addWidget(QLabel(str(int(u.get("total") or 0))), i, 1)
            tlay.addWidget(QLabel(str(int(u.get("correct") or 0))), i, 2)
            tlay.addWidget(QLabel(f"{float(u.get('rate_pct') or 0):.1f}%"), i, 3)
        lay.addWidget(table_wrap)
    lay.addStretch(1)
    return wrap


class StudentReportTabWidget(QWidget):
    """학생 페이지 > 보고서 탭: 저장된 보고서 목록 + 생성/상세 뷰."""

    def __init__(
        self,
        db_connection,
        *,
        student_id: str,
        student_name: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db_connection = db_connection
        self.student_id = (student_id or "").strip()
        self.student_name = (student_name or "").strip()

        self._inner_stack: Optional[QStackedWidget] = None
        self._list_layout: Optional[QVBoxLayout] = None
        self._list_scroll: Optional[QScrollArea] = None
        self._create_form: Optional[QWidget] = None
        self._detail_widget: Optional[QWidget] = None
        self._current_snapshot: Optional[dict] = None
        self._current_report_id: Optional[str] = None
        self._build_ui()

    def _build_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self._inner_stack = QStackedWidget()

        # 0: 목록 뷰 (1단 구조: 헤더 + 풀폭 리스트)
        list_page = QWidget()
        list_page.setStyleSheet("QWidget { background: #FFFFFF; }")
        list_page_lay = QVBoxLayout(list_page)
        list_page_lay.setContentsMargins(0, 0, 0, 0)
        list_page_lay.setSpacing(0)

        # 헤더: 제목 + 우측 '새 보고서 생성' 버튼
        header = QWidget()
        header.setStyleSheet("QWidget { background: #FFFFFF; border: none; }")
        header_lay = QHBoxLayout(header)
        header_lay.setContentsMargins(20, 16, 20, 16)
        header_lay.setSpacing(0)
        lbl_title = QLabel("보고서 관리")
        lbl_title.setStyleSheet("background: none; color: #1A1A1A; font-size: 14pt; font-weight: bold;")
        lbl_title.setFont(_font(14, bold=True))
        header_lay.addWidget(lbl_title)
        header_lay.addStretch(1)
        btn_create = QPushButton("+ 새 보고서 생성")
        btn_create.setObjectName("ReportCreateBtn")
        btn_create.setFont(_font(11, bold=True))
        btn_create.setFixedHeight(36)
        btn_create.setCursor(Qt.PointingHandCursor)
        btn_create.setStyleSheet(
            """
            QPushButton#ReportCreateBtn {
                background: #2563EB;
                color: #FFFFFF;
                border: none;
                border-radius: 8px;
                padding: 0 16px;
                font-size: 11pt;
            }
            QPushButton#ReportCreateBtn:hover {
                background: #1D4ED8;
            }
            """
        )
        btn_create.clicked.connect(self._go_create)
        header_lay.addWidget(btn_create)
        list_page_lay.addWidget(header)

        # 구분선 (1px)
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("background: #E5E7EB; max-height: 1px; border: none;")
        list_page_lay.addWidget(sep)

        # 리스트 영역 (풀폭)
        self._list_scroll = QScrollArea()
        self._list_scroll.setWidgetResizable(True)
        self._list_scroll.setFrameShape(QFrame.NoFrame)
        self._list_scroll.setStyleSheet("QScrollArea { background: #FFFFFF; border: none; }")
        list_content = QWidget()
        list_content.setStyleSheet("QWidget { background: #FFFFFF; }")
        self._list_layout = QVBoxLayout(list_content)
        self._list_layout.setContentsMargins(0, 12, 0, 0)
        self._list_layout.setSpacing(0)
        self._list_scroll.setWidget(list_content)
        list_page_lay.addWidget(self._list_scroll, 1)
        self._inner_stack.addWidget(list_page)

        # 1: 생성/상세 뷰 (공용 컨테이너, 내용 교체, 배경 순백)
        self._form_container = QWidget()
        self._form_container.setStyleSheet("background: #FFFFFF;")
        self._form_container.setLayout(QVBoxLayout())
        self._inner_stack.addWidget(self._form_container)

        root.addWidget(self._inner_stack, 1)
        self._show_list()

    def _show_list(self) -> None:
        self._inner_stack.setCurrentIndex(0)
        self._reload_list()

    def _reload_list(self) -> None:
        if self._list_layout is None:
            return
        while self._list_layout.count():
            it = self._list_layout.takeAt(0)
            w = it.widget()
            if w:
                w.setParent(None)

        if not self.db_connection or not self.db_connection.is_connected() or not self.student_id:
            empty = QLabel("저장된 보고서를 불러올 수 없습니다.")
            empty.setStyleSheet("color: #64748B; font-size: 11pt;")
            self._list_layout.addWidget(empty)
            return

        try:
            repo = ReportRepository(self.db_connection)
            reports = repo.list_by_student(self.student_id)
        except Exception:
            reports = []

        if not reports:
            empty = QLabel("저장된 보고서가 없습니다. 우측 상단 '+ 새 보고서 생성'으로 새 보고서를 만드세요.")
            empty.setStyleSheet("color: #64748B; font-size: 11pt;")
            empty.setWordWrap(True)
            self._list_layout.addWidget(empty)
            return

        for r in reports:
            row = _ReportListRow(r, self.student_name)
            row.clicked.connect(self._open_detail)
            row.edit_requested.connect(self._on_edit_report)
            row.delete_requested.connect(self._on_delete_report)
            self._list_layout.addWidget(row)
        self._list_layout.addStretch(1)

    def _go_create(self) -> None:
        if not self.db_connection or not self.db_connection.is_connected() or not self.student_id:
            show_warning(self, "보고서", "DB 연결 또는 학생 정보가 없습니다.")
            return
        modal = ReportCreateModal(self.db_connection, self.student_id, self.student_name, self)
        if modal.exec_() == QDialog.Accepted:
            self._reload_list()

    def _open_detail(self, report_id: str) -> None:
        if not self.db_connection or not self.db_connection.is_connected():
            return
        try:
            repo = ReportRepository(self.db_connection)
            report = repo.get_by_id(report_id)
        except Exception:
            report = None
        if not report:
            show_warning(self, "보고서", "보고서를 찾을 수 없습니다.")
            return
        modal = ReportDetailModal(report, self.student_name, self)
        modal.exec_()

    def _on_edit_report(self, report_id: str) -> None:
        """수정 클릭 시 학습코멘트만 편집 가능한 다이얼로그."""
        if not report_id or not self.db_connection or not self.db_connection.is_connected():
            return
        try:
            repo = ReportRepository(self.db_connection)
            report = repo.get_by_id(report_id)
        except Exception:
            report = None
        if not report:
            show_warning(self, "수정", "보고서를 찾을 수 없습니다.")
            return
        current = (report.comment or "").strip()
        modal = ReportCommentEditDialog(self, current)
        if modal.exec_() != QDialog.Accepted:
            return
        text = modal.get_comment()
        try:
            ok = repo.update_comment(report_id, text)
        except Exception:
            ok = False
        if ok:
            show_info(self, "수정 완료", "학습코멘트가 저장되었습니다.")
            self._reload_list()
        else:
            show_warning(self, "수정 실패", "학습코멘트 저장에 실패했습니다.")

    def _on_delete_report(self, report_id: str) -> None:
        if not report_id:
            return
        try:
            repo = ReportRepository(self.db_connection)
            ok = repo.delete(report_id)
        except Exception:
            ok = False
        if ok:
            show_info(self, "삭제", "보고서가 삭제되었습니다.")
            if self._current_report_id == report_id:
                self._show_list()
            else:
                self._reload_list()
        else:
            show_warning(self, "삭제 실패", "보고서를 삭제하지 못했습니다.")


class StudentPage(QWidget):
    """학생별 관리 페이지(학습지/오답노트/보고서 탭)."""

    def __init__(
        self,
        db_connection,
        *,
        student_id: str,
        student_name: str,
        student_grade: str,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.db_connection = db_connection
        self.student_id = (student_id or "").strip()
        self.student_name = (student_name or "").strip()
        self.student_grade = (student_grade or "").strip()

        self._stack: Optional[QStackedWidget] = None
        self._build_ui()

    def _build_ui(self) -> None:
        self.setObjectName("StudentPageRoot")
        self.setAutoFillBackground(True)
        pal = self.palette()
        pal.setColor(QPalette.Window, QColor(0xFF, 0xFF, 0xFF))
        self.setPalette(pal)
        self.setStyleSheet(
            """
            QWidget#StudentPageRoot {
                background-color: #FFFFFF;
            }
            /* 수업 탭 내 학생 페이지: 학습지 리스트 등 회색 배경 제거 → 흰색 통일 */
            QWidget#StudentPageRoot QStackedWidget,
            QWidget#StudentPageRoot QWidget#worksheetList {
                background-color: #FFFFFF;
            }

            QLabel#StudentName {
                color: #0F172A;
            }
            QLabel#StudentMeta {
                color: #64748B;
            }

            /* 탭: 테두리 없음, 클릭한 탭에만 밑줄 */
            QPushButton#StudentTabBtn {
                color: #64748B;
                font-size: 12pt;
                font-weight: 700;
                padding: 10px 12px;
                padding-bottom: 12px;
                border: none;
                border-bottom: 3px solid transparent;
                border-radius: 0px;
                background: transparent;
                outline: none;
            }
            QPushButton#StudentTabBtn:hover {
                color: #2563EB;
            }
            QPushButton#StudentTabBtn[active="true"] {
                color: #2563EB;
                font-weight: 900;
                border: none;
                border-bottom: 3px solid #2563EB;
            }
            """
        )

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        # 상단: 학생명 + (학년)
        top = QHBoxLayout()
        top.setContentsMargins(0, 0, 0, 0)
        top.setSpacing(10)

        name = QLabel(self.student_name or "학생")
        name.setObjectName("StudentName")
        name.setFont(_font(14, extra_bold=True))
        top.addWidget(name, alignment=Qt.AlignVCenter)

        meta = QLabel(self.student_grade)
        meta.setObjectName("StudentMeta")
        meta.setFont(_font(10, bold=True))
        top.addWidget(meta, alignment=Qt.AlignVCenter)

        top.addStretch(1)
        root.addLayout(top)

        # 탭 바
        tab_row = QHBoxLayout()
        tab_row.setContentsMargins(0, 0, 0, 0)
        tab_row.setSpacing(18)

        self.btn_ws = QPushButton("학습지")
        self.btn_wrong = QPushButton("오답노트")
        self.btn_report = QPushButton("보고서")
        for b in (self.btn_ws, self.btn_wrong, self.btn_report):
            b.setObjectName("StudentTabBtn")
            b.setCheckable(True)
            b.setFocusPolicy(Qt.NoFocus)
            b.setCursor(Qt.PointingHandCursor)
            b.setProperty("active", False)
            b.toggled.connect(lambda checked, btn=b: self._sync_tab(btn, checked))
            tab_row.addWidget(b)

        tab_row.addStretch(1)
        root.addLayout(tab_row)

        # 컨텐츠
        self._stack = QStackedWidget()
        root.addWidget(self._stack, 1)

        # 1) 학습지 탭(출제된 것만 + 채점/오답노트 생성)
        self._ws_screen = StudentWorksheetListScreen(
            self.db_connection,
            student_id=self.student_id,
            student_name=self.student_name,
            student_grade=self.student_grade,
        )
        self._ws_screen.grading_saved.connect(self._on_grading_saved)
        self._ws_screen.wrongnote_ready.connect(self._open_wrongnote_tab)
        self._stack.addWidget(self._ws_screen)

        # 2) 오답노트 탭(생성된 오답노트만 표시)
        self._wrong_screen = StudentWrongNoteListScreen(
            self.db_connection,
            student_id=self.student_id,
            student_name=self.student_name,
        )
        self._stack.addWidget(self._wrong_screen)

        # 3) 보고서 탭(목록 + 생성/상세)
        self._report_screen = StudentReportTabWidget(
            self.db_connection,
            student_id=self.student_id,
            student_name=self.student_name,
        )
        self._stack.addWidget(self._report_screen)

        # 기본: 학습지 (선택된 탭에만 밑줄 표시)
        self.btn_ws.setChecked(True)
        self.btn_ws.setProperty("active", True)
        for o in (self.btn_wrong, self.btn_report):
            o.setChecked(False)
            o.setProperty("active", False)
        self._stack.setCurrentIndex(0)
        self._refresh_tab_style()

        self.btn_ws.clicked.connect(lambda: self._stack.setCurrentIndex(0))
        self.btn_wrong.clicked.connect(lambda: self._stack.setCurrentIndex(1))
        self.btn_report.clicked.connect(lambda: self._stack.setCurrentIndex(2))

    def _open_wrongnote_tab(self, worksheet_id: str = "") -> None:
        # 오답노트 탭으로 이동 + 목록 갱신
        try:
            if getattr(self, "_wrong_screen", None) is not None:
                self._wrong_screen.reload_from_db()
        except Exception:
            pass
        try:
            self.btn_wrong.setChecked(True)
            self._stack.setCurrentIndex(1)
        except Exception:
            pass

    def _on_grading_saved(self, worksheet_id: str) -> None:
        # 채점 변경 시 오답노트 목록도 최신화
        try:
            if getattr(self, "_wrong_screen", None) is not None:
                self._wrong_screen.reload_from_db()
        except Exception:
            pass

    def _sync_tab(self, btn: QPushButton, checked: bool) -> None:
        if checked:
            for o in (self.btn_ws, self.btn_wrong, self.btn_report):
                if o is not btn:
                    o.setChecked(False)
                    o.setProperty("active", False)
            btn.setProperty("active", True)
        else:
            btn.setProperty("active", False)
        self._refresh_tab_style()

    def _refresh_tab_style(self) -> None:
        for btn in (self.btn_ws, self.btn_wrong, self.btn_report):
            try:
                btn.style().unpolish(btn)
                btn.style().polish(btn)
                btn.update()
            except Exception:
                pass


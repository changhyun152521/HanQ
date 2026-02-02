"""
학습 보고서 집계 서비스

- 기간 내 채점 완료된 출제 건을 조회하여 요약·단원별 통계를 계산합니다.
- 교재/기출별 분석은 문제(Problem) 단위 출처(source_id, source_type)로 집계합니다.
  학습지(Worksheet) type_text/title은 문제 출처가 없을 때만 fallback으로 사용합니다.
"""

from __future__ import annotations

from typing import Any, Dict, List

from core.models import SourceType
from database.sqlite_connection import SQLiteConnection
from database.repositories.worksheet_assignment_repository import WorksheetAssignmentRepository
from database.repositories.worksheet_repository import WorksheetRepository
from database.repositories.problem_repository import ProblemRepository
from database.repositories.textbook_repository import TextbookRepository
from database.repositories.exam_repository import ExamRepository


def _exam_display_name(exam) -> str:
    """기출 표시명: 예) OO학교 O학년 O학기 중간고사 2026"""
    parts = [
        getattr(exam, "school_name", None) or "",
        getattr(exam, "grade", None) or "",
        getattr(exam, "semester", None) or "",
        getattr(exam, "exam_type", None) or "",
        getattr(exam, "year", None) or "",
    ]
    name = " ".join(str(p).strip() for p in parts if p).strip()
    return name or "미명"


def aggregate_report(
    db_connection: SQLiteConnection,
    student_id: str,
    period_start: str,
    period_end: str,
) -> Dict[str, Any]:
    """
    학생의 기간 내 채점 완료 데이터를 집계하여 보고서 스냅샷을 반환합니다.

    Args:
        db_connection: DB 연결
        student_id: 학생 ID
        period_start: 시작일 (YYYY-MM-DD)
        period_end: 종료일 (YYYY-MM-DD)

    Returns:
        snapshot dict:
        - total_worksheets: int
        - total_questions: int
        - total_correct: int
        - average_rate_pct: float
        - unit_stats: [{"unit_key", "unit_label", "total", "correct", "rate_pct"}, ...]
    """
    sid = (student_id or "").strip()
    start_s = (period_start or "").strip()
    end_s = (period_end or "").strip()
    if not sid or not start_s or not end_s:
        return _empty_snapshot()

    if not db_connection.is_connected():
        return _empty_snapshot()

    repo = WorksheetAssignmentRepository(db_connection)
    docs = repo.list_graded_for_student_in_period(sid, start_s, end_s)
    if not docs:
        return _empty_snapshot()

    total_worksheets = len(docs)
    total_questions = 0
    total_correct = 0
    unit_merged: Dict[str, Dict[str, int]] = {}  # unit_key -> {total, correct}
    source_merged: Dict[tuple, Dict[str, int]] = {}  # (category, name) -> {total, correct}

    ws_repo = WorksheetRepository(db_connection)
    problem_repo = ProblemRepository(db_connection)
    textbook_repo = TextbookRepository(db_connection)
    exam_repo = ExamRepository(db_connection)

    for doc in docs:
        tq = int(doc.get("total_questions") or 0)
        tc = int(doc.get("correct_count") or 0)
        total_questions += tq
        total_correct += tc
        ustats = doc.get("unit_stats") or {}
        if isinstance(ustats, dict):
            for uk, uv in ustats.items():
                if not isinstance(uv, dict):
                    continue
                ut = int(uv.get("total") or 0)
                uc = int(uv.get("correct") or 0)
                if uk not in unit_merged:
                    unit_merged[uk] = {"total": 0, "correct": 0}
                unit_merged[uk]["total"] += ut
                unit_merged[uk]["correct"] += uc

        # 교재/기출별: 문제(Problem) 단위 출처로 집계 (출처 데이터가 있는 문항만)
        added_from_problems = False
        raw_answers = doc.get("answers")
        answers = raw_answers if isinstance(raw_answers, list) else []
        for ans in answers:
            if not isinstance(ans, dict):
                continue
            try:
                pid = str(ans.get("problem_id") or "").strip()
                if not pid:
                    continue
                is_correct = bool(ans.get("is_correct"))
                problem = problem_repo.find_by_id(pid) if problem_repo else None
                if not problem:
                    continue
                raw_source_id = getattr(problem, "source_id", None)
                source_id = str(raw_source_id or "").strip()
                if not source_id:
                    continue
                source_type = getattr(problem, "source_type", None)
                category = "기출" if source_type == SourceType.EXAM else "교재"
                name = ""
                try:
                    if source_type == SourceType.EXAM:
                        exam = exam_repo.find_by_id(source_id) if exam_repo else None
                        if exam:
                            name = _exam_display_name(exam)
                    else:
                        textbook = textbook_repo.find_by_id(source_id) if textbook_repo else None
                        if textbook:
                            name = str(getattr(textbook, "name", None) or "").strip()
                except Exception:
                    pass
                if not name:
                    name = "미명"
                key = (category, name)
                if key not in source_merged:
                    source_merged[key] = {"total": 0, "correct": 0}
                source_merged[key]["total"] += 1
                source_merged[key]["correct"] += 1 if is_correct else 0
                added_from_problems = True
            except Exception:
                continue

        # 문제 출처가 하나도 없으면 학습지(Worksheet) 단위으로 1건 fallback
        if not added_from_problems:
            wid = str(doc.get("worksheet_id") or "").strip()
            ws = ws_repo.find_by_id(wid) if wid else None
            type_text = (getattr(ws, "type_text", None) or "") if ws else ""
            category = "기출" if "기출" in type_text else "교재"
            name = (getattr(ws, "title", None) or "").strip() if ws else "미확인"
            if not name:
                name = "미명"
            key = (category, name)
            if key not in source_merged:
                source_merged[key] = {"total": 0, "correct": 0}
            source_merged[key]["total"] += tq
            source_merged[key]["correct"] += tc

    average_rate_pct = (total_correct / total_questions * 100.0) if total_questions else 0.0
    unit_stats: List[Dict[str, Any]] = []
    for uk, uv in sorted(unit_merged.items(), key=lambda x: (-(x[1]["total"]), x[0])):
        t = uv["total"]
        c = uv["correct"]
        rate = (c / t * 100.0) if t else 0.0
        unit_stats.append({
            "unit_key": uk,
            "unit_label": uk or "미분류",
            "total": t,
            "correct": c,
            "rate_pct": round(rate, 1),
        })

    textbook_stats: List[Dict[str, Any]] = []
    exam_stats: List[Dict[str, Any]] = []
    for (category, name), uv in sorted(source_merged.items(), key=lambda x: (-(x[1]["total"]), x[0][0], x[0][1])):
        t = uv["total"]
        c = uv["correct"]
        item = {"name": name, "correct": c, "total": t}
        if category == "기출":
            exam_stats.append(item)
        else:
            textbook_stats.append(item)

    return {
        "total_worksheets": total_worksheets,
        "total_questions": total_questions,
        "total_correct": total_correct,
        "average_rate_pct": round(average_rate_pct, 1),
        "unit_stats": unit_stats,
        "textbook_stats": textbook_stats,
        "exam_stats": exam_stats,
    }


def _empty_snapshot() -> Dict[str, Any]:
    return {
        "total_worksheets": 0,
        "total_questions": 0,
        "total_correct": 0,
        "average_rate_pct": 0.0,
        "unit_stats": [],
        "textbook_stats": [],
        "exam_stats": [],
    }

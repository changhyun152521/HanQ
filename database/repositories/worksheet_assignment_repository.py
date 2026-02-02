"""
WorksheetAssignment 저장소

- SQLite 테이블: worksheet_assignments
"""
from __future__ import annotations

import json
from datetime import datetime
from typing import Dict, Iterable, List, Optional

from database.sqlite_connection import SQLiteConnection, json_col


class WorksheetAssignmentRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection

    @staticmethod
    def _now_iso() -> str:
        return datetime.now().isoformat()

    def assign_many(
        self,
        *,
        worksheet_ids: Iterable[str],
        student_ids: Iterable[str],
        assigned_by: str = "",
    ) -> dict:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        ws_ids = [str(x).strip() for x in (worksheet_ids or []) if str(x).strip()]
        st_ids = [str(x).strip() for x in (student_ids or []) if str(x).strip()]
        if not ws_ids or not st_ids:
            return {"inserted": 0, "skipped": 0, "total": 0}

        now = self._now_iso()
        conn = self._db.get_conn()
        inserted = 0
        for wid in ws_ids:
            for sid in st_ids:
                try:
                    cur = conn.execute(
                        """INSERT INTO worksheet_assignments (
                            worksheet_id, student_id, assigned_at, assigned_by, status
                        ) VALUES (?, ?, ?, ?, ?)
                        ON CONFLICT(worksheet_id, student_id) DO NOTHING""",
                        (wid, sid, now, (assigned_by or "").strip(), "assigned"),
                    )
                    if cur.rowcount > 0:
                        inserted += 1
                except Exception:
                    pass
        conn.commit()
        total = len(ws_ids) * len(st_ids)
        return {"inserted": inserted, "skipped": total - inserted, "total": total}

    def list_for_student(self, student_id: str) -> List[dict]:
        sid = (student_id or "").strip()
        if not sid:
            return []
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM worksheet_assignments WHERE student_id = ? ORDER BY assigned_at DESC",
                (sid,),
            ).fetchall()
            return [_assignment_row_to_doc(r) for r in rows]
        except Exception:
            return []

    def list_graded_for_student_in_period(
        self,
        student_id: str,
        period_start: str,
        period_end: str,
    ) -> List[dict]:
        sid = (student_id or "").strip()
        start_s = (period_start or "").strip()
        end_s = (period_end or "").strip()
        if not sid or not start_s or not end_s:
            return []
        start_iso = start_s if "T" in start_s else f"{start_s}T00:00:00"
        end_iso = end_s if "T" in end_s else f"{end_s}T23:59:59.999999"
        try:
            rows = self._db.get_conn().execute(
                """SELECT * FROM worksheet_assignments
                   WHERE student_id = ? AND status = 'graded'
                     AND graded_at >= ? AND graded_at <= ?
                   ORDER BY graded_at DESC""",
                (sid, start_iso, end_iso),
            ).fetchall()
            return [_assignment_row_to_doc(r) for r in rows]
        except Exception:
            return []

    def list_wrongnotes_for_student(self, student_id: str) -> List[dict]:
        sid = (student_id or "").strip()
        if not sid:
            return []
        try:
            rows = self._db.get_conn().execute(
                """SELECT * FROM worksheet_assignments
                   WHERE student_id = ? AND wrongnote_enabled = 1
                   ORDER BY assigned_at DESC""",
                (sid,),
            ).fetchall()
            return [_assignment_row_to_doc(r) for r in rows]
        except Exception:
            return []

    def enable_wrongnote(
        self,
        *,
        worksheet_id: str,
        student_id: str,
        title: str,
    ) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        wid = (worksheet_id or "").strip()
        sid = (student_id or "").strip()
        if not wid or not sid:
            return False
        try:
            cur = self._db.get_conn().execute(
                """UPDATE worksheet_assignments SET
                    wrongnote_enabled = 1, wrongnote_title = ?, wrongnote_updated_at = ?
                WHERE worksheet_id = ? AND student_id = ?""",
                ((title or "").strip(), self._now_iso(), wid, sid),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def set_wrong_info(
        self,
        *,
        worksheet_id: str,
        student_id: str,
        wrong_problem_ids: List[str],
    ) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        wid = (worksheet_id or "").strip()
        sid = (student_id or "").strip()
        if not wid or not sid:
            return False
        ids = [str(x).strip() for x in (wrong_problem_ids or []) if str(x).strip()]
        try:
            cur = self._db.get_conn().execute(
                """UPDATE worksheet_assignments SET
                    wrong_problem_ids_json = ?, wrong_count = ?, wrongnote_updated_at = ?
                WHERE worksheet_id = ? AND student_id = ?""",
                (json_col(ids, "[]"), len(ids), self._now_iso(), wid, sid),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def find_one(
        self,
        *,
        worksheet_id: str,
        student_id: str,
    ) -> Optional[dict]:
        wid = (worksheet_id or "").strip()
        sid = (student_id or "").strip()
        if not wid or not sid:
            return None
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM worksheet_assignments WHERE worksheet_id = ? AND student_id = ?",
                (wid, sid),
            ).fetchone()
            if not row:
                return None
            return _assignment_row_to_doc(row)
        except Exception:
            return None

    def save_grading(
        self,
        *,
        worksheet_id: str,
        student_id: str,
        total_questions: int,
        correct_count: int,
        answers: List[Dict],
        unit_stats: Dict[str, Dict],
        assigned_by: str = "",
    ) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        wid = (worksheet_id or "").strip()
        sid = (student_id or "").strip()
        if not wid or not sid:
            return False
        wrong_ids = []
        for a in (answers or []):
            try:
                pid = str(a.get("problem_id") or "").strip()
                if pid and not bool(a.get("is_correct")):
                    wrong_ids.append(pid)
            except Exception:
                pass
        now = self._now_iso()
        try:
            cur = self._db.get_conn().execute(
                """UPDATE worksheet_assignments SET
                    status = 'graded', graded_at = ?, total_questions = ?, correct_count = ?,
                    answers_json = ?, unit_stats_json = ?, wrong_problem_ids_json = ?,
                    wrong_count = ?, wrongnote_updated_at = ?, assigned_by = ?
                WHERE worksheet_id = ? AND student_id = ?""",
                (
                    now,
                    int(total_questions),
                    int(correct_count),
                    json_col(answers, "[]"),
                    json_col(unit_stats, "{}"),
                    json_col(wrong_ids, "[]"),
                    len(wrong_ids),
                    now,
                    (assigned_by or "").strip(),
                    wid,
                    sid,
                ),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def save_wrongnote_grading(
        self,
        *,
        worksheet_id: str,
        student_id: str,
        total_questions: int,
        correct_count: int,
        answers: List[Dict],
        unit_stats: Dict[str, Dict],
        assigned_by: str = "",
    ) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        wid = (worksheet_id or "").strip()
        sid = (student_id or "").strip()
        if not wid or not sid:
            return False
        now = self._now_iso()
        try:
            cur = self._db.get_conn().execute(
                """UPDATE worksheet_assignments SET
                    wrongnote_status = 'graded', wrongnote_graded_at = ?,
                    wrongnote_total_questions = ?, wrongnote_correct_count = ?,
                    wrongnote_answers_json = ?, wrongnote_unit_stats_json = ?,
                    wrongnote_updated_at = ?, assigned_by = ?
                WHERE worksheet_id = ? AND student_id = ?""",
                (
                    now,
                    int(total_questions),
                    int(correct_count),
                    json_col(answers, "[]"),
                    json_col(unit_stats, "{}"),
                    now,
                    (assigned_by or "").strip(),
                    wid,
                    sid,
                ),
            )
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False


def _assignment_row_to_doc(row) -> dict:
    d = dict(row)
    doc = {
        "_id": str(d["id"]),
        "worksheet_id": str(d.get("worksheet_id", "")),
        "student_id": str(d.get("student_id", "")),
        "assigned_at": d.get("assigned_at"),
        "assigned_by": d.get("assigned_by") or "",
        "status": d.get("status") or "assigned",
        "graded_at": d.get("graded_at"),
        "score": d.get("score"),
        "total_questions": d.get("total_questions"),
        "correct_count": d.get("correct_count"),
        "answers": _parse_json(d.get("answers_json"), []),
        "unit_stats": _parse_json(d.get("unit_stats_json"), {}),
        "wrong_problem_ids": _parse_json(d.get("wrong_problem_ids_json"), []),
        "wrong_count": d.get("wrong_count"),
        "wrongnote_enabled": bool(d.get("wrongnote_enabled")),
        "wrongnote_title": d.get("wrongnote_title"),
        "wrongnote_updated_at": d.get("wrongnote_updated_at"),
        "wrongnote_status": d.get("wrongnote_status"),
        "wrongnote_graded_at": d.get("wrongnote_graded_at"),
        "wrongnote_total_questions": d.get("wrongnote_total_questions"),
        "wrongnote_correct_count": d.get("wrongnote_correct_count"),
        "wrongnote_answers": _parse_json(d.get("wrongnote_answers_json"), []),
        "wrongnote_unit_stats": _parse_json(d.get("wrongnote_unit_stats_json"), {}),
    }
    return doc


def _parse_json(s, default):
    if not s:
        return default
    if isinstance(s, str):
        try:
            return json.loads(s)
        except Exception:
            return default
    return default if s is None else s

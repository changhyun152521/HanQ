"""
Problem 저장소

- SQLite 테이블: problems
- HWP 바이너리: file_store
"""
from __future__ import annotations

import json
from typing import List, Optional

from core.models import Problem, SourceType
from database.sqlite_connection import SQLiteConnection, row_to_dict, json_col


def _parse_json(s, default):
    if not s:
        return default
    if isinstance(s, str):
        try:
            return json.loads(s)
        except Exception:
            return default
    return default if s is None else s


class ProblemRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection
        self._store = db_connection.get_file_store()

    def create(self, problem: Problem, hwp_bytes: bytes) -> str:
        if not self._db.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 저장할 수 없습니다."
            )
        file_id = self._store.put(hwp_bytes, kind="application/x-hwp", filename="problem.hwp")
        problem.content_raw_file_id = file_id
        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO problems (
                content_raw_file_id, content_text, source_id, source_type,
                tags_json, created_at, creator, original_hwp_path, problem_index
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                int(file_id),
                problem.content_text or "",
                problem.source_id or "",
                problem.source_type.value,
                json_col([t.to_dict() for t in (problem.tags or [])]),
                problem.created_at.isoformat() if problem.created_at else None,
                problem.creator or "",
                problem.original_hwp_path,
                problem.problem_index or 0,
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def find_by_id(self, problem_id: str) -> Optional[Problem]:
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM problems WHERE id = ?", (int(problem_id),)
            ).fetchone()
            if not row:
                return None
            d = row_to_dict(row)
            d["content_raw_file_id"] = str(d["content_raw_file_id"]) if d.get("content_raw_file_id") else None
            d["tags"] = _parse_json(d.get("tags_json"), [])
            return Problem.from_dict(d)
        except (ValueError, TypeError):
            return None

    def get_content_raw(self, problem_id: str) -> Optional[bytes]:
        try:
            p = self.find_by_id(problem_id)
            if p and p.content_raw_file_id:
                return self._store.get(p.content_raw_file_id)
            return None
        except Exception:
            return None

    def find_by_source(self, source_id: str, source_type: SourceType) -> List[Problem]:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM problems WHERE source_id = ? AND source_type = ?",
                (source_id, source_type.value),
            ).fetchall()
            out = []
            for row in rows:
                d = row_to_dict(row)
                d["content_raw_file_id"] = str(d["content_raw_file_id"]) if d.get("content_raw_file_id") else None
                d["tags"] = _parse_json(d.get("tags_json"), [])
                out.append(Problem.from_dict(d))
            return out
        except Exception:
            return []

    def search_by_text(self, keyword: str) -> List[Problem]:
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM problems WHERE content_text LIKE ?",
                (f"%{keyword}%",),
            ).fetchall()
            out = []
            for row in rows:
                d = row_to_dict(row)
                d["content_raw_file_id"] = str(d["content_raw_file_id"]) if d.get("content_raw_file_id") else None
                d["tags"] = _parse_json(d.get("tags_json"), [])
                out.append(Problem.from_dict(d))
            return out
        except Exception:
            return []

    def update(self, problem: Problem) -> bool:
        try:
            if not problem.id:
                return False
            conn = self._db.get_conn()
            cur = conn.execute(
                """UPDATE problems SET
                    content_text = ?, source_id = ?, source_type = ?,
                    tags_json = ?, creator = ?, original_hwp_path = ?, problem_index = ?
                WHERE id = ?""",
                (
                    problem.content_text or "",
                    problem.source_id or "",
                    problem.source_type.value,
                    json_col([t.to_dict() for t in (problem.tags or [])]),
                    problem.creator or "",
                    problem.original_hwp_path,
                    problem.problem_index or 0,
                    int(problem.id),
                ),
            )
            conn.commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def delete(self, problem_id: str) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        try:
            p = self.find_by_id(problem_id)
            if p and p.content_raw_file_id:
                self._store.delete(p.content_raw_file_id)
            cur = self._db.get_conn().execute("DELETE FROM problems WHERE id = ?", (int(problem_id),))
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False

    def batch_create(self, problems: List[tuple]) -> List[str]:
        ids = []
        for problem, hwp_bytes in problems:
            try:
                pid = self.create(problem, hwp_bytes)
                ids.append(pid)
            except Exception:
                pass
        return ids

    def list_by_ids(self, problem_ids: List[str]) -> List[Problem]:
        ids = [str(x).strip() for x in (problem_ids or []) if str(x).strip()]
        if not ids:
            return []
        try:
            placeholders = ",".join("?" * len(ids))
            int_ids = [int(x) for x in ids]
            rows = self._db.get_conn().execute(
                f"SELECT * FROM problems WHERE id IN ({placeholders})",
                int_ids,
            ).fetchall()
            by_id = {}
            for row in rows:
                d = row_to_dict(row)
                d["content_raw_file_id"] = str(d["content_raw_file_id"]) if d.get("content_raw_file_id") else None
                d["tags"] = _parse_json(d.get("tags_json"), [])
                by_id[str(d["_id"])] = Problem.from_dict(d)
            return [by_id[pid] for pid in ids if pid in by_id]
        except Exception:
            return []

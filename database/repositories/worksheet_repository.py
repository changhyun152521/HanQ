"""
Worksheet 저장소

- SQLite 테이블: worksheets
- HWP/PDF 바이너리: file_store
"""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from core.models import Worksheet
from database.sqlite_connection import SQLiteConnection, row_to_dict, json_col


class WorksheetRepository:
    def __init__(self, db_connection: SQLiteConnection):
        self._db = db_connection
        self._store = db_connection.get_file_store()

    def create(
        self,
        worksheet: Worksheet,
        *,
        hwp_bytes: bytes,
        pdf_bytes: Optional[bytes] = None,
        hwp_filename: Optional[str] = None,
        pdf_filename: Optional[str] = None,
    ) -> str:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        if worksheet.created_at is None:
            worksheet.created_at = datetime.now()

        hwp_file_id = self._store.put(
            hwp_bytes,
            kind="application/x-hwp",
            filename=hwp_filename or "worksheet.hwp",
        )
        worksheet.hwp_file_id = hwp_file_id

        pdf_file_id = None
        if pdf_bytes:
            pdf_file_id = self._store.put(
                pdf_bytes,
                kind="application/pdf",
                filename=pdf_filename or "worksheet.pdf",
            )
            worksheet.pdf_file_id = pdf_file_id

        conn = self._db.get_conn()
        conn.execute(
            """INSERT INTO worksheets (
                title, grade, type_text, creator, created_at,
                problem_ids_json, numbered_json, hwp_file_id, pdf_file_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
            (
                worksheet.title or "",
                worksheet.grade or "",
                worksheet.type_text or "",
                worksheet.creator or "",
                worksheet.created_at.isoformat() if worksheet.created_at else None,
                json_col(worksheet.problem_ids, "[]"),
                json_col(worksheet.numbered, "[]"),
                int(hwp_file_id),
                int(pdf_file_id) if pdf_file_id else None,
            ),
        )
        conn.commit()
        return str(conn.execute("SELECT last_insert_rowid()").fetchone()[0])

    def find_by_id(self, worksheet_id: str) -> Optional[Worksheet]:
        try:
            row = self._db.get_conn().execute(
                "SELECT * FROM worksheets WHERE id = ?", (int(worksheet_id),)
            ).fetchone()
            if not row:
                return None
            d = row_to_dict(row)
            d["hwp_file_id"] = str(d["hwp_file_id"]) if d.get("hwp_file_id") else None
            d["pdf_file_id"] = str(d["pdf_file_id"]) if d.get("pdf_file_id") else None
            d["problem_ids"] = _parse_json(d.get("problem_ids_json"), [])
            d["numbered"] = _parse_json(d.get("numbered_json"), [])
            return Worksheet.from_dict(d)
        except (ValueError, TypeError):
            return None

    def list_all(self) -> List[Worksheet]:
        try:
            rows = self._db.get_conn().execute(
                "SELECT * FROM worksheets ORDER BY created_at DESC"
            ).fetchall()
            out = []
            for row in rows:
                d = row_to_dict(row)
                d["hwp_file_id"] = str(d["hwp_file_id"]) if d.get("hwp_file_id") else None
                d["pdf_file_id"] = str(d["pdf_file_id"]) if d.get("pdf_file_id") else None
                d["problem_ids"] = _parse_json(d.get("problem_ids_json"), [])
                d["numbered"] = _parse_json(d.get("numbered_json"), [])
                out.append(Worksheet.from_dict(d))
            return out
        except Exception:
            return []

    def list_by_ids(self, worksheet_ids: List[str]) -> List[Worksheet]:
        ids = [str(x).strip() for x in (worksheet_ids or []) if str(x).strip()]
        if not ids:
            return []
        try:
            placeholders = ",".join("?" * len(ids))
            int_ids = [int(x) for x in ids]
            rows = self._db.get_conn().execute(
                f"SELECT * FROM worksheets WHERE id IN ({placeholders})",
                int_ids,
            ).fetchall()
            by_id = {}
            for row in rows:
                d = row_to_dict(row)
                d["hwp_file_id"] = str(d["hwp_file_id"]) if d.get("hwp_file_id") else None
                d["pdf_file_id"] = str(d["pdf_file_id"]) if d.get("pdf_file_id") else None
                d["problem_ids"] = _parse_json(d.get("problem_ids_json"), [])
                d["numbered"] = _parse_json(d.get("numbered_json"), [])
                by_id[str(d["_id"])] = Worksheet.from_dict(d)
            return [by_id[wid] for wid in ids if wid in by_id]
        except Exception:
            return []

    def get_file_bytes(self, worksheet: Worksheet, kind: str) -> Optional[bytes]:
        k = (kind or "").strip().upper()
        file_id_str = worksheet.hwp_file_id if k == "HWP" else worksheet.pdf_file_id
        if not file_id_str:
            return None
        return self._store.get(file_id_str)

    def delete(self, worksheet_id: str) -> bool:
        if not self._db.is_connected():
            raise ConnectionError("DB에 연결되지 않았습니다.")
        try:
            ws = self.find_by_id(worksheet_id)
            if ws:
                for fid in [ws.hwp_file_id, ws.pdf_file_id]:
                    if fid:
                        self._store.delete(fid)
            cur = self._db.get_conn().execute("DELETE FROM worksheets WHERE id = ?", (int(worksheet_id),))
            self._db.get_conn().commit()
            return cur.rowcount > 0
        except Exception:
            return False


def _parse_json(s, default):
    if not s:
        return default
    if isinstance(s, str):
        try:
            import json
            return json.loads(s)
        except Exception:
            return default
    return default if s is None else s

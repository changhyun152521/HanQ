"""
SQLite 연결 및 파일 저장소

- 단일 DB 파일로 데이터 관리 (배포·이동 시 파일만 복사)
- 테이블 기반 스키마 + file_store(HWP/PDF 바이너리)
- 변수·이름: 연결(connection), 파일저장소(file_store), id만 사용
"""

from __future__ import annotations

import json
import os
import sqlite3
from typing import Optional


class FileStore:
    """HWP/PDF 등 바이너리 저장·조회 (GridFS 대체)."""

    def __init__(self, conn: sqlite3.Connection):
        self._conn = conn

    def put(self, data: bytes, kind: str = "", filename: str = "") -> str:
        """바이트 저장 후 id 반환."""
        cur = self._conn.execute(
            "INSERT INTO file_store (kind, filename, data) VALUES (?, ?, ?)",
            (kind or "", filename or "", data),
        )
        self._conn.commit()
        return str(cur.lastrowid)

    def get(self, file_id: str) -> Optional[bytes]:
        """id로 바이트 조회."""
        try:
            row = self._conn.execute(
                "SELECT data FROM file_store WHERE id = ?",
                (int(file_id),),
            ).fetchone()
            return row[0] if row else None
        except (ValueError, TypeError):
            return None

    def delete(self, file_id: str) -> bool:
        """id로 삭제."""
        try:
            cur = self._conn.execute("DELETE FROM file_store WHERE id = ?", (int(file_id),))
            self._conn.commit()
            return cur.rowcount > 0
        except (ValueError, TypeError):
            return False


def _schema_sql() -> str:
    return """
CREATE TABLE IF NOT EXISTS file_store (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    kind TEXT NOT NULL DEFAULT '',
    filename TEXT NOT NULL DEFAULT '',
    data BLOB NOT NULL
);

CREATE TABLE IF NOT EXISTS problems (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    content_raw_file_id INTEGER,
    content_text TEXT NOT NULL DEFAULT '',
    source_id TEXT NOT NULL DEFAULT '',
    source_type TEXT NOT NULL DEFAULT 'textbook',
    tags_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT,
    creator TEXT NOT NULL DEFAULT '',
    original_hwp_path TEXT,
    problem_index INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS textbooks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL DEFAULT '',
    subject TEXT NOT NULL DEFAULT '',
    major_unit TEXT NOT NULL DEFAULT '',
    sub_unit TEXT,
    created_at TEXT,
    parsed_at TEXT,
    is_parsed INTEGER NOT NULL DEFAULT 0,
    problem_count INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS exams (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    grade TEXT NOT NULL DEFAULT '',
    semester TEXT NOT NULL DEFAULT '',
    exam_type TEXT NOT NULL DEFAULT '',
    school_name TEXT NOT NULL DEFAULT '',
    year TEXT NOT NULL DEFAULT '',
    created_at TEXT,
    parsed_at TEXT,
    is_parsed INTEGER NOT NULL DEFAULT 0,
    problem_count INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS worksheets (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL DEFAULT '',
    grade TEXT NOT NULL DEFAULT '',
    type_text TEXT NOT NULL DEFAULT '',
    creator TEXT NOT NULL DEFAULT '',
    created_at TEXT,
    problem_ids_json TEXT NOT NULL DEFAULT '[]',
    numbered_json TEXT NOT NULL DEFAULT '[]',
    hwp_file_id INTEGER,
    pdf_file_id INTEGER
);

CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    grade TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT '재원',
    name TEXT NOT NULL DEFAULT '',
    school_name TEXT NOT NULL DEFAULT '',
    parent_phone TEXT NOT NULL DEFAULT '',
    student_phone TEXT NOT NULL DEFAULT '',
    created_at TEXT,
    updated_at TEXT,
    deleted_at TEXT
);

CREATE TABLE IF NOT EXISTS classes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    grade TEXT NOT NULL DEFAULT '',
    name TEXT NOT NULL DEFAULT '',
    teacher TEXT NOT NULL DEFAULT '',
    note TEXT NOT NULL DEFAULT '',
    student_ids_json TEXT NOT NULL DEFAULT '[]',
    created_at TEXT,
    updated_at TEXT,
    deleted_at TEXT
);

CREATE TABLE IF NOT EXISTS worksheet_assignments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    worksheet_id TEXT NOT NULL,
    student_id TEXT NOT NULL,
    assigned_at TEXT NOT NULL,
    assigned_by TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'assigned',
    graded_at TEXT,
    score REAL,
    total_questions INTEGER,
    correct_count INTEGER,
    answers_json TEXT,
    unit_stats_json TEXT,
    wrong_problem_ids_json TEXT,
    wrong_count INTEGER,
    wrongnote_enabled INTEGER NOT NULL DEFAULT 0,
    wrongnote_title TEXT,
    wrongnote_updated_at TEXT,
    wrongnote_status TEXT,
    wrongnote_graded_at TEXT,
    wrongnote_total_questions INTEGER,
    wrongnote_correct_count INTEGER,
    wrongnote_answers_json TEXT,
    wrongnote_unit_stats_json TEXT,
    UNIQUE(worksheet_id, student_id)
);

CREATE INDEX IF NOT EXISTS ix_wa_student_assigned ON worksheet_assignments(student_id, assigned_at);
CREATE INDEX IF NOT EXISTS ix_wa_worksheet_assigned ON worksheet_assignments(worksheet_id, assigned_at);

CREATE TABLE IF NOT EXISTS saved_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id TEXT NOT NULL,
    period_start TEXT NOT NULL DEFAULT '',
    period_end TEXT NOT NULL DEFAULT '',
    comment TEXT NOT NULL DEFAULT '',
    created_at TEXT,
    snapshot_json TEXT NOT NULL DEFAULT '{}'
);

CREATE INDEX IF NOT EXISTS ix_saved_reports_student_created ON saved_reports(student_id, created_at);
"""


class SQLiteConnection:
    """SQLite 단일 파일 연결. is_connected / get_conn / get_file_store 만 노출."""

    def __init__(self, db_path: str):
        self._path = os.path.abspath(db_path)
        self._conn: Optional[sqlite3.Connection] = None
        self._file_store: Optional[FileStore] = None

    def connect(self) -> bool:
        """DB 파일 생성·연결 및 스키마 초기화."""
        try:
            parent = os.path.dirname(self._path)
            if parent:
                os.makedirs(parent, exist_ok=True)
            self._conn = sqlite3.connect(self._path)
            self._conn.row_factory = sqlite3.Row
            self._conn.executescript(_schema_sql())
            self._conn.commit()
            self._file_store = FileStore(self._conn)
            return True
        except Exception:
            return False

    def disconnect(self) -> None:
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None
        self._file_store = None

    def is_connected(self) -> bool:
        if self._conn is None:
            return False
        try:
            self._conn.execute("SELECT 1").fetchone()
            return True
        except Exception:
            return False

    def get_conn(self) -> sqlite3.Connection:
        if self._conn is None:
            raise RuntimeError("DB에 연결되지 않았습니다.")
        return self._conn

    def get_file_store(self) -> FileStore:
        if self._file_store is None:
            raise RuntimeError("DB에 연결되지 않았습니다.")
        return self._file_store


def _parse_dt(s: Optional[str]):
    if not s:
        return None
    if isinstance(s, str):
        try:
            from datetime import datetime
            return datetime.fromisoformat(s.replace("Z", "+00:00"))
        except Exception:
            return None
    return s


def row_to_dict(row: sqlite3.Row, *, id_key: str = "_id") -> dict:
    """SQLite Row를 dict로. id 컬럼을 id_key(기본 _id)로 넣어 모델 호환."""
    d = dict(row)
    if "id" in d and id_key != "id":
        d[id_key] = str(d["id"])
        del d["id"]
    return d


def json_col(val, default: str = "[]") -> str:
    if val is None:
        return default
    if isinstance(val, str):
        return val
    return json.dumps(val, ensure_ascii=False)

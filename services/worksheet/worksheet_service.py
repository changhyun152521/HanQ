"""
WorksheetService (1단계: 문제 선택 엔진 연결)

현재 단계에서는 "문제 선택"까지만 제공합니다.
- UI에서 선택된 Textbook.id / Exam.id
- UI에서 선택된 단원(UnitKey 리스트)
- 총 문항수/난이도 비율/정렬 옵션

을 받아, DB에서 후보 문제를 모아서 selection_engine으로 최종 문제 id 리스트를 반환합니다.

향후 단계:
- 선택된 문제들을 HWP 문서로 합치기(Composer)
- Worksheet 메타데이터/파일(GridFS) 저장
- worksheet_list.py와 연결(다운로드/삭제/목록)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple

from core.models import Problem, SourceType
from database.sqlite_connection import SQLiteConnection
from database.repositories import ProblemRepository

from .selection_engine import (
    DifficultyRatios,
    OrderOptions,
    UnitKey,
    WorksheetSelectionEngine,
    WorksheetSelectionResult,
    WorksheetSelectionSpec,
)


@dataclass(frozen=True)
class SelectedSources:
    """
    UI에서 선택된 출처 집합
    - textbook_ids: 선택된 Textbook.id들
    - exam_ids: 선택된 Exam.id들
    """

    textbook_ids: List[str]
    exam_ids: List[str]

    def is_empty(self) -> bool:
        return not (self.textbook_ids or self.exam_ids)


class WorksheetService:
    """학습지 관련 서비스(1단계: 문제 선택)."""

    def __init__(self, db_connection: SQLiteConnection):
        self.db_connection = db_connection
        self.problem_repo = ProblemRepository(db_connection)
        self.engine = WorksheetSelectionEngine()

    def _fetch_candidates_by_sources(self, sources: SelectedSources) -> List[Problem]:
        """
        선택된 출처에서 Problem 후보를 모두 로드합니다.
        - 현재 Repository API는 source_id 단위 조회만 제공 → 여러 source_id를 순회
        """
        if sources.is_empty():
            return []

        problems: List[Problem] = []
        seen: set[str] = set()

        for tid in sources.textbook_ids:
            for p in self.problem_repo.find_by_source(tid, SourceType.TEXTBOOK):
                if p.id and p.id not in seen:
                    seen.add(p.id)
                    problems.append(p)

        for eid in sources.exam_ids:
            for p in self.problem_repo.find_by_source(eid, SourceType.EXAM):
                if p.id and p.id not in seen:
                    seen.add(p.id)
                    problems.append(p)

        return problems

    def select_problems(
        self,
        *,
        units: List[UnitKey],
        sources: SelectedSources,
        total_count: int,
        difficulty_ratios: DifficultyRatios,
        order: OrderOptions,
        seed: Optional[int] = None,
    ) -> WorksheetSelectionResult:
        """
        1단계: 문제 선택(문제 id 리스트 반환)

        예외/정책:
        - 출처는 최소 1개(Textbook 또는 Exam)라도 선택되어야 함
        - 단원은 1개 이상 선택되어야 함
        - 단원 중 어떤 단원이 0문항이어도 생성은 계속(총 문항 감소)
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 조회할 수 없습니다."
            )

        if sources.is_empty():
            raise ValueError("출처(교재 또는 내신기출)를 1개 이상 선택해야 합니다.")

        candidates = self._fetch_candidates_by_sources(sources)
        spec = WorksheetSelectionSpec(
            units=units,
            total_count=int(total_count),
            difficulty_ratios=difficulty_ratios,
            order=order,
            seed=seed,
        )
        return self.engine.select(spec, candidates)


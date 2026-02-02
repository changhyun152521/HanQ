"""
학습지(Worksheet) 서비스 모듈

- 문제 선택 엔진(단원별 균등배분 + 난이도 비율 + 부족 시 감산)
- 향후: HWP 문서 합치기/저장/다운로드까지 확장
"""

from .selection_engine import (
    DifficultyRatios,
    OrderOptions,
    UnitKey,
    WorksheetSelectionEngine,
    WorksheetSelectionResult,
    WorksheetSelectionSpec,
)
from .worksheet_service import WorksheetService

__all__ = [
    "UnitKey",
    "DifficultyRatios",
    "OrderOptions",
    "WorksheetSelectionSpec",
    "WorksheetSelectionResult",
    "WorksheetSelectionEngine",
    "WorksheetService",
]


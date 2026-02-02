"""
학습 보고서(Report) 서비스 모듈

- 기간별 채점 데이터 집계(요약·단원별 정답률)
- 저장된 보고서 스냅샷 생성
"""

from .report_service import aggregate_report

__all__ = ["aggregate_report"]

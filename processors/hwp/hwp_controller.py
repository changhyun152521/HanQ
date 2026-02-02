# -*- coding: utf-8 -*-
"""
HWP 문서 처리 컨트롤러

HWP 문서 처리의 전체 흐름을 관리하는 컨트롤러입니다.
- HWP 파일 선택 및 열기
- ProblemSplitter와 HWPReader 조율
- Problem 객체 생성 및 저장
- 처리 진행 상황 관리
- 에러 처리 및 사용자 피드백
"""
import json
import os
from typing import List, Optional
from datetime import datetime
from core.models import Problem, SourceType, Tag
from processors.hwp.problem_splitter import ProblemSplitter
from database.sqlite_connection import SQLiteConnection
from database.repositories import ProblemRepository, TextbookRepository, ExamRepository


def _default_db_path() -> str:
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    path = "./db/ch_lms.db"
    try:
        cfg_path = os.path.join(root, "config", "config.json")
        if os.path.isfile(cfg_path):
            with open(cfg_path, "r", encoding="utf-8") as f:
                path = (json.load(f).get("database") or {}).get("path") or path
    except Exception:
        pass
    return path if os.path.isabs(path) else os.path.join(root, path)


class HWPController:
    """HWP 파싱 컨트롤러"""

    def __init__(self, db_connection: Optional[SQLiteConnection] = None):
        """
        HWPController 초기화

        Args:
            db_connection: DB 연결 (None이면 기본 경로로 자동 생성)
        """
        self.splitter = ProblemSplitter()
        if db_connection:
            self.db_connection = db_connection
        else:
            self.db_connection = SQLiteConnection(_default_db_path())
            self.db_connection.connect()

        self.problem_repo = ProblemRepository(self.db_connection)
        self.textbook_repo = TextbookRepository(self.db_connection)
        self.exam_repo = ExamRepository(self.db_connection)
    
    def parse_hwp_to_problems(
        self,
        hwp_path: str,
        source_id: str,
        source_type: SourceType,
        creator: str = "",
        progress_callback: Optional[callable] = None,
        *,
        apply_style_to_blocks: bool = False
    ) -> List[Problem]:
        """
        HWP 파일을 파싱하여 Problem 리스트 생성 및 저장
        
        Args:
            hwp_path: HWP 파일 경로
            source_id: Textbook.id 또는 Exam.id
            source_type: SourceType Enum
            creator: 출제자명
            progress_callback: 진행 상황 콜백 함수 (current, total) -> None
            apply_style_to_blocks: True면 저장되는 각 문제 문서에 스타일 적용. 기본 False(파싱만, 잘림 방지).
            
        Returns:
            생성된 Problem 객체 리스트
        """
        problems = []
        
        try:
            # 교재DB는 파일 1개 = 단원 1개인 경우가 많아, 파싱 시점에 Problem에 단원 태그를 자동 부여합니다.
            textbook_tag_template = None
            if source_type == SourceType.TEXTBOOK:
                try:
                    textbook = self.textbook_repo.find_by_id(source_id)
                    if textbook:
                        unit_str = textbook.major_unit
                        if textbook.sub_unit:
                            unit_str = f"{textbook.major_unit} > {textbook.sub_unit}"
                        textbook_tag_template = {
                            "subject": textbook.subject,
                            "major_unit": textbook.major_unit,
                            "sub_unit": textbook.sub_unit,
                            "unit": unit_str,
                        }
                except Exception:
                    textbook_tag_template = None

            # 문제 블록 추출 (미주 기반 파싱)
            print(f"[디버그] HWP 파싱 시작: {hwp_path}")
            problem_blocks = self.splitter.split_problems_by_endnote(
                hwp_path, apply_style_to_blocks=apply_style_to_blocks
            )
            total = len(problem_blocks)
            
            print(f"[디버그] 추출된 문제 블록 수: {total}")
            
            if total == 0:
                print("[경고] 추출된 문제가 없습니다. HWP 파일에 미주가 올바르게 있는지 확인하세요.")
                return problems
            
            # 진행 상황 콜백 호출
            if progress_callback:
                progress_callback(0, total)
            
            # 각 문제 블록마다 Problem 생성 및 저장
            for index, (hwp_bytes, text) in enumerate(problem_blocks, start=1):
                try:
                    tags = []
                    if textbook_tag_template:
                        tags = [
                            Tag(
                                subject=textbook_tag_template["subject"],
                                grade="",
                                major_unit=textbook_tag_template["major_unit"],
                                sub_unit=textbook_tag_template["sub_unit"],
                                unit=textbook_tag_template["unit"],
                                difficulty=None,
                            )
                        ]
                    # Problem 객체 생성
                    problem = Problem(
                        content_text=text,
                        source_id=source_id,
                        source_type=source_type,
                        tags=tags,
                        created_at=datetime.now(),
                        creator=creator,
                        original_hwp_path=hwp_path,
                        problem_index=index
                    )
                    
                    # MongoDB에 저장 (GridFS 포함)
                    problem_id = self.problem_repo.create(problem, hwp_bytes)
                    problem.id = problem_id
                    
                    problems.append(problem)
                    
                    # 진행 상황 콜백 호출
                    if progress_callback:
                        progress_callback(index, total)
                
                except Exception as e:
                    print(f"문제 {index} 저장 실패: {e}")
                    continue
            
            # 출처 메타데이터 업데이트
            self._update_source_parsed_status(source_id, source_type, len(problems))
            
            return problems
        
        except Exception as e:
            print(f"HWP 파싱 실패: {e}")
            return problems
    
    def _update_source_parsed_status(
        self,
        source_id: str,
        source_type: SourceType,
        problem_count: int
    ):
        """
        출처 메타데이터의 파싱 상태 업데이트
        
        Args:
            source_id: Textbook.id 또는 Exam.id
            source_type: SourceType Enum
            problem_count: 생성된 Problem 개수
        """
        try:
            if source_type == SourceType.TEXTBOOK:
                self.textbook_repo.update_parsed_status(
                    source_id,
                    is_parsed=True,
                    problem_count=problem_count
                )
            elif source_type == SourceType.EXAM:
                self.exam_repo.update_parsed_status(
                    source_id,
                    is_parsed=True,
                    problem_count=problem_count
                )
        except Exception as e:
            print(f"출처 메타데이터 업데이트 실패: {e}")
    
    def cleanup(self):
        """리소스 정리"""
        if self.db_connection and not self.db_connection.is_connected():
            self.db_connection.disconnect()

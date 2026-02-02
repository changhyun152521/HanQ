"""
파싱 재실행 Service

Textbook/Exam 단위로 HWP 파싱을 재실행하는 서비스를 제공합니다.
- 기존 Problem 삭제 후 재파싱
- 기존 Problem 유지 + 신규 Problem 추가
- 파싱 상태 및 메타데이터 업데이트
"""
from typing import Optional, List, Dict, Any
from enum import Enum
from database.repositories import ProblemRepository, TextbookRepository, ExamRepository
from database.sqlite_connection import SQLiteConnection
from processors.hwp.hwp_controller import HWPController
from processors.hwp.hwp_reader import HWPNotInstalledError, HWPInitializationError
from core.models import Problem, SourceType, Textbook, Exam


class ReparseMode(Enum):
    """재파싱 모드"""
    REPLACE = "replace"  # 기존 Problem 전부 삭제 후 재파싱
    APPEND = "append"     # 기존 Problem 유지 + 신규 Problem 추가


class ParsingService:
    """파싱 재실행 서비스"""
    
    def __init__(self, db_connection: SQLiteConnection):
        """
        ParsingService 초기화

        Args:
            db_connection: DB 연결 인스턴스
        """
        self.db_connection = db_connection
        self.problem_repo = ProblemRepository(db_connection)
        self.textbook_repo = TextbookRepository(db_connection)
        self.exam_repo = ExamRepository(db_connection)
        self.hwp_controller = HWPController(db_connection)
    
    def reparse_textbook(
        self,
        textbook_id: str,
        hwp_path: str,
        mode: ReparseMode = ReparseMode.REPLACE,
        creator: str = "",
        progress_callback: Optional[callable] = None,
        *,
        apply_style_to_blocks: bool = False
    ) -> Dict[str, Any]:
        """
        Textbook 단위로 HWP 파싱 재실행
        
        Args:
            textbook_id: Textbook ID
            hwp_path: HWP 파일 경로
            mode: 재파싱 모드 (REPLACE 또는 APPEND)
            creator: 출제자명
            progress_callback: 진행 상황 콜백 함수 (current, total) -> None
            apply_style_to_blocks: True면 저장되는 각 문제 문서에 스타일 적용(본문/미주/주석 줄간격 등). 기본 False.
            
        Returns:
            {
                'success': bool,
                'deleted_count': int,
                'created_count': int,
                'total_problems': int,
                'error': str or None
            }
            
        Raises:
            HWPNotInstalledError: 한글 프로그램이 설치되지 않은 경우
            HWPInitializationError: 한글 프로그램 초기화 실패 시
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        # Textbook 조회
        textbook = self.textbook_repo.find_by_id(textbook_id)
        if not textbook:
            return {
                'success': False,
                'deleted_count': 0,
                'created_count': 0,
                'total_problems': 0,
                'error': f'Textbook을 찾을 수 없습니다. (ID: {textbook_id})'
            }
        
        # 기존 Problem 삭제 (REPLACE 모드인 경우)
        deleted_count = 0
        if mode == ReparseMode.REPLACE:
            existing_problems = self.problem_repo.find_by_source(
                textbook_id,
                SourceType.TEXTBOOK
            )
            for problem in existing_problems:
                if problem.id:
                    self.problem_repo.delete(problem.id)
                    deleted_count += 1
            
            # 파싱 상태 초기화
            self.textbook_repo.update_parsed_status(
                textbook_id,
                is_parsed=False,
                problem_count=0
            )
        
        # HWP 파싱 실행
        try:
            problems = self.hwp_controller.parse_hwp_to_problems(
                hwp_path=hwp_path,
                source_id=textbook_id,
                source_type=SourceType.TEXTBOOK,
                creator=creator,
                progress_callback=progress_callback,
                apply_style_to_blocks=apply_style_to_blocks
            )
            
            created_count = len(problems)
            total_problems = self.problem_repo.find_by_source(
                textbook_id,
                SourceType.TEXTBOOK
            )
            
            return {
                'success': True,
                'deleted_count': deleted_count,
                'created_count': created_count,
                'total_problems': len(total_problems),
                'error': None
            }
        
        except (HWPNotInstalledError, HWPInitializationError) as e:
            return {
                'success': False,
                'deleted_count': deleted_count,
                'created_count': 0,
                'total_problems': 0,
                'error': str(e)
            }
        except Exception as e:
            return {
                'success': False,
                'deleted_count': deleted_count,
                'created_count': 0,
                'total_problems': 0,
                'error': f'파싱 중 오류 발생: {str(e)}'
            }
    
    def reparse_exam(
        self,
        exam_id: str,
        hwp_path: str,
        mode: ReparseMode = ReparseMode.REPLACE,
        creator: str = "",
        progress_callback: Optional[callable] = None,
        *,
        apply_style_to_blocks: bool = False
    ) -> Dict[str, Any]:
        """
        Exam 단위로 HWP 파싱 재실행
        
        Args:
            exam_id: Exam ID
            hwp_path: HWP 파일 경로
            mode: 재파싱 모드 (REPLACE 또는 APPEND)
            creator: 출제자명
            progress_callback: 진행 상황 콜백 함수 (current, total) -> None
            apply_style_to_blocks: True면 저장되는 각 문제 문서에 스타일 적용. 기본 False.
            
        Returns:
            {
                'success': bool,
                'deleted_count': int,
                'created_count': int,
                'total_problems': int,
                'error': str or None
            }
            
        Raises:
            HWPNotInstalledError: 한글 프로그램이 설치되지 않은 경우
            HWPInitializationError: 한글 프로그램 초기화 실패 시
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        # Exam 조회
        exam = self.exam_repo.find_by_id(exam_id)
        if not exam:
            return {
                'success': False,
                'deleted_count': 0,
                'created_count': 0,
                'total_problems': 0,
                'error': f'Exam을 찾을 수 없습니다. (ID: {exam_id})'
            }
        
        # 기존 Problem 삭제 (REPLACE 모드인 경우)
        deleted_count = 0
        if mode == ReparseMode.REPLACE:
            existing_problems = self.problem_repo.find_by_source(
                exam_id,
                SourceType.EXAM
            )
            for problem in existing_problems:
                if problem.id:
                    self.problem_repo.delete(problem.id)
                    deleted_count += 1
            
            # 파싱 상태 초기화
            self.exam_repo.update_parsed_status(
                exam_id,
                is_parsed=False,
                problem_count=0
            )
        
        # HWP 파싱 실행
        try:
            problems = self.hwp_controller.parse_hwp_to_problems(
                hwp_path=hwp_path,
                source_id=exam_id,
                source_type=SourceType.EXAM,
                creator=creator,
                progress_callback=progress_callback,
                apply_style_to_blocks=apply_style_to_blocks
            )
            
            created_count = len(problems)
            total_problems = self.problem_repo.find_by_source(
                exam_id,
                SourceType.EXAM
            )
            
            return {
                'success': True,
                'deleted_count': deleted_count,
                'created_count': created_count,
                'total_problems': len(total_problems),
                'error': None
            }
        
        except (HWPNotInstalledError, HWPInitializationError) as e:
            return {
                'success': False,
                'deleted_count': deleted_count,
                'created_count': 0,
                'total_problems': 0,
                'error': str(e)
            }
        except Exception as e:
            return {
                'success': False,
                'deleted_count': deleted_count,
                'created_count': 0,
                'total_problems': 0,
                'error': f'파싱 중 오류 발생: {str(e)}'
            }

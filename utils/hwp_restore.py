"""
원본 HWP 복원 유틸리티

GridFS에 저장된 Problem의 content_raw_file_id를 이용해
원본 HWP를 임시 파일(.hwp)로 복원하는 기능을 제공합니다.
- UI에서 "원본 HWP 보기" 버튼에 바로 연결 가능
"""
import os
import tempfile
from typing import Optional, List
from database.repositories import ProblemRepository
from database.sqlite_connection import SQLiteConnection


class HWPRestoreError(Exception):
    """HWP 복원 중 오류 발생 시 예외"""
    pass


class HWPRestore:
    """원본 HWP 복원 유틸리티"""
    
    def __init__(self, db_connection: SQLiteConnection):
        """
        HWPRestore 초기화

        Args:
            db_connection: DB 연결 인스턴스
        """
        self.db_connection = db_connection
        self.problem_repo = ProblemRepository(db_connection)
    
    def restore_to_file(
        self,
        problem_id: str,
        output_path: Optional[str] = None,
        temp_dir: Optional[str] = None
    ) -> str:
        """
        Problem의 원본 HWP를 파일로 복원
        
        Args:
            problem_id: Problem ID
            output_path: 저장할 파일 경로 (None이면 임시 파일 생성)
            temp_dir: 임시 파일 저장 디렉토리 (output_path가 None일 때 사용)
            
        Returns:
            복원된 HWP 파일 경로
            
        Raises:
            HWPRestoreError: 복원 실패 시
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. HWP를 복원할 수 없습니다."
            )
        
        # Problem 조회
        problem = self.problem_repo.find_by_id(problem_id)
        if not problem:
            raise HWPRestoreError(f"Problem을 찾을 수 없습니다. (ID: {problem_id})")
        
        if not problem.content_raw_file_id:
            raise HWPRestoreError(f"Problem에 원본 HWP가 저장되어 있지 않습니다. (ID: {problem_id})")
        
        # file_store에서 HWP 바이너리 읽기
        try:
            hwp_bytes = self.problem_repo.get_content_raw(problem_id)
            if not hwp_bytes:
                raise HWPRestoreError(f"HWP 바이너리를 읽을 수 없습니다. (ID: {problem_id})")
        except Exception as e:
            raise HWPRestoreError(f"HWP 바이너리 조회 실패: {str(e)}")
        
        # 출력 파일 경로 결정
        if output_path is None:
            if temp_dir is None:
                temp_dir = tempfile.gettempdir()
            
            # 임시 파일명 생성 (problem_id 기반)
            filename = f"problem_{problem_id}.hwp"
            output_path = os.path.join(temp_dir, filename)
        
        # 파일 저장
        try:
            with open(output_path, 'wb') as f:
                f.write(hwp_bytes)
            return output_path
        except Exception as e:
            raise HWPRestoreError(f"HWP 파일 저장 실패: {str(e)}")
    
    def restore_to_temp_file(
        self,
        problem_id: str,
        prefix: str = "problem_",
        suffix: str = ".hwp"
    ) -> str:
        """
        Problem의 원본 HWP를 임시 파일로 복원
        
        Args:
            problem_id: Problem ID
            prefix: 임시 파일명 접두사
            suffix: 임시 파일명 접미사
            
        Returns:
            복원된 임시 HWP 파일 경로
            
        Raises:
            HWPRestoreError: 복원 실패 시
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        # 임시 파일 생성
        temp_fd, temp_path = tempfile.mkstemp(prefix=prefix, suffix=suffix)
        os.close(temp_fd)  # 파일 디스크립터 닫기 (파일은 유지)
        
        try:
            # 복원 실행
            return self.restore_to_file(problem_id, output_path=temp_path)
        except Exception as e:
            # 실패 시 임시 파일 삭제
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except:
                    pass
            raise
    
    def restore_multiple_to_dir(
        self,
        problem_ids: List[str],
        output_dir: str
    ) -> List[str]:
        """
        여러 Problem의 원본 HWP를 한 디렉토리에 복원
        
        Args:
            problem_ids: Problem ID 리스트
            output_dir: 저장할 디렉토리 경로
            
        Returns:
            복원된 HWP 파일 경로 리스트
            
        Raises:
            HWPRestoreError: 복원 실패 시
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        restored_paths = []
        for problem_id in problem_ids:
            try:
                # 파일명 생성 (problem_id 기반)
                filename = f"problem_{problem_id}.hwp"
                output_path = os.path.join(output_dir, filename)
                
                # 복원 실행
                restored_path = self.restore_to_file(problem_id, output_path=output_path)
                restored_paths.append(restored_path)
            except Exception as e:
                print(f"Problem {problem_id} 복원 실패: {e}")
                continue
        
        return restored_paths

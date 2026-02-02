"""
의존성 체크 유틸리티

프로그램 실행에 필요한 외부 의존성을 확인합니다.
- 한글 프로그램 설치 여부 확인
- MongoDB 연결 가능 여부 확인 (선택적)
"""
import sys
from typing import Tuple, Optional


class DependencyChecker:
    """의존성 체크 클래스"""
    
    @staticmethod
    def check_hwp_installed() -> Tuple[bool, Optional[str]]:
        """
        한글 프로그램 설치 여부 확인
        
        Returns:
            (설치 여부, 에러 메시지)
        """
        if sys.platform != 'win32':
            return False, "이 프로그램은 Windows 환경에서만 동작합니다."
        
        try:
            import win32com.client
            hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
            hwp.Quit()
            return True, None
        except Exception as e:
            error_msg = (
                "한글 프로그램이 설치되어 있지 않거나 COM 자동화가 불가능합니다.\n"
                "한글과컴퓨터의 한글 프로그램을 설치한 후 다시 시도해주세요.\n"
                f"상세 오류: {str(e)}"
            )
            return False, error_msg
    
    @staticmethod
    def check_mongodb_connection(host: str = "localhost", port: int = 27017, timeout: int = 3) -> Tuple[bool, Optional[str]]:
        """
        MongoDB 연결 가능 여부 확인 (선택적)
        
        Args:
            host: MongoDB 호스트
            port: MongoDB 포트
            timeout: 연결 타임아웃 (초)
            
        Returns:
            (연결 가능 여부, 에러 메시지)
        """
        try:
            from pymongo import MongoClient
            from pymongo.errors import ServerSelectionTimeoutError, ConnectionFailure
            
            client = MongoClient(host, port, serverSelectionTimeoutMS=timeout * 1000)
            client.admin.command('ping')
            client.close()
            return True, None
        except (ServerSelectionTimeoutError, ConnectionFailure) as e:
            # 연결 실패는 정상 시나리오로 처리
            return False, f"MongoDB 서버에 연결할 수 없습니다. (호스트: {host}:{port})"
        except Exception as e:
            return False, f"MongoDB 연결 확인 중 오류 발생: {str(e)}"
    
    @staticmethod
    def check_all_dependencies() -> Tuple[bool, list]:
        """
        모든 필수 의존성 확인
        
        Returns:
            (모든 의존성 충족 여부, 실패한 의존성 목록)
        """
        failures = []
        
        # 한글 프로그램 체크 (필수)
        hwp_ok, hwp_error = DependencyChecker.check_hwp_installed()
        if not hwp_ok:
            failures.append(("한글 프로그램", hwp_error))
        
        return len(failures) == 0, failures
    
    @staticmethod
    def get_dependency_status_message() -> str:
        """
        의존성 상태 메시지 생성
        
        Returns:
            상태 메시지 문자열
        """
        messages = []
        
        # 한글 프로그램 체크
        hwp_ok, hwp_error = DependencyChecker.check_hwp_installed()
        if hwp_ok:
            messages.append("✓ 한글 프로그램: 설치됨")
        else:
            messages.append(f"✗ 한글 프로그램: {hwp_error}")
        
        # MongoDB 체크 (선택적, 연결 실패는 경고만)
        mongodb_ok, mongodb_error = DependencyChecker.check_mongodb_connection()
        if mongodb_ok:
            messages.append("✓ MongoDB: 연결 가능")
        else:
            messages.append(f"⚠ MongoDB: {mongodb_error} (오프라인 모드로 동작 가능)")
        
        return "\n".join(messages)

"""
MongoDB 연결 및 설정

MongoDB 데이터베이스 연결을 관리합니다.
- MongoDB 클라이언트 생성
- 데이터베이스 및 컬렉션 선택
- GridFS 설정
- 연결 상태 확인
"""
from pymongo import MongoClient
from gridfs import GridFS
from typing import Optional
import json
import os


class MongoDBConnection:
    """MongoDB 연결 관리 클래스"""
    
    def __init__(self, config_path: Optional[str] = None):
        """
        MongoDB 연결 초기화
        
        Args:
            config_path: config.json 파일 경로 (기본값: config/config.json)
        """
        if config_path is None:
            config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config', 'config.json')
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        mongodb_config = config.get('mongodb', {})
        self.host = mongodb_config.get('host', 'localhost')
        self.port = mongodb_config.get('port', 27017)
        self.database_name = mongodb_config.get('database', 'ch_lms')
        
        self.client: Optional[MongoClient] = None
        self.db = None
        self.fs: Optional[GridFS] = None  # GridFS 인스턴스
    
    def connect(self):
        """MongoDB 연결"""
        try:
            self.client = MongoClient(self.host, self.port)
            self.db = self.client[self.database_name]
            self.fs = GridFS(self.db)  # GridFS 초기화
            
            # 연결 테스트
            self.client.admin.command('ping')
            return True
        except Exception as e:
            print(f"MongoDB 연결 실패: {e}")
            return False
    
    def disconnect(self):
        """MongoDB 연결 종료"""
        if self.client:
            self.client.close()
            self.client = None
            self.db = None
            self.fs = None
    
    def is_connected(self) -> bool:
        """연결 상태 확인"""
        if self.client is None:
            return False
        try:
            self.client.admin.command('ping')
            return True
        except:
            return False
    
    def get_collection(self, collection_name: str):
        """컬렉션 가져오기"""
        if self.db is None:
            raise Exception("MongoDB에 연결되지 않았습니다.")
        return self.db[collection_name]
    
    def get_gridfs(self) -> GridFS:
        """GridFS 인스턴스 가져오기"""
        if self.fs is None:
            raise Exception("MongoDB에 연결되지 않았습니다.")
        return self.fs
    
    def __enter__(self):
        """Context manager 진입"""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager 종료"""
        self.disconnect()


# 전역 연결 인스턴스 (선택적 사용)
_global_connection: Optional[MongoDBConnection] = None


def get_connection() -> MongoDBConnection:
    """전역 MongoDB 연결 인스턴스 가져오기"""
    global _global_connection
    if _global_connection is None:
        _global_connection = MongoDBConnection()
        _global_connection.connect()
    return _global_connection
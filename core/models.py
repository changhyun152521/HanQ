"""
데이터 모델 정의

MongoDB에 저장될 문제 데이터의 구조를 정의합니다.
- Problem 모델 클래스 (HWP 원본 보존)
- Tag 모델 클래스
- Textbook, Exam 메타데이터 모델
- SourceType Enum
"""
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from datetime import datetime
from enum import Enum
import re


class SourceType(Enum):
    """출처 타입 Enum"""
    TEXTBOOK = "textbook"
    EXAM = "exam"


@dataclass
class Tag:
    """문제 태그 모델"""
    subject: str  # 과목 (예: 수학, 영어)
    grade: str    # 학년 (예: 초5, 중1, 고2)
    # 단원 (기출DB는 파일 1개에 여러 단원이 섞일 수 있어 Problem별로 태깅)
    major_unit: Optional[str] = None  # 대단원
    sub_unit: Optional[str] = None  # 소단원
    unit: Optional[str] = None  # (레거시) 단원명 문자열. major/sub를 채울 때 함께 채움
    difficulty: Optional[str] = None  # 난이도 (킬, 상, 중, 하)
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환"""
        return {
            'subject': self.subject,
            'grade': self.grade,
            'major_unit': self.major_unit,
            'sub_unit': self.sub_unit,
            'unit': self.unit,
            'difficulty': self.difficulty
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Tag':
        """딕셔너리에서 생성"""
        major_unit = data.get('major_unit')
        sub_unit = data.get('sub_unit')
        unit = data.get('unit')

        # 과거 데이터(major/sub 없음)는 unit 문자열에서 최대한 복원
        if (not major_unit and not sub_unit) and isinstance(unit, str) and unit.strip():
            parts = [p.strip() for p in re.split(r"\s*>\s*|\s*/\s*", unit) if p.strip()]
            if parts:
                major_unit = parts[0] or None
                if len(parts) > 1:
                    sub_unit = parts[1] or None

        return cls(
            subject=data.get('subject', ''),
            grade=data.get('grade', ''),
            major_unit=major_unit,
            sub_unit=sub_unit,
            unit=unit,
            difficulty=data.get('difficulty')
        )


@dataclass
class Problem:
    """문제 모델 - HWP 원본 블록을 그대로 보존"""
    id: Optional[str] = None  # MongoDB ObjectId
    
    # HWP 원본 데이터 (최우선)
    content_raw_file_id: Optional[str] = None  # GridFS 파일 ID
    content_text: str = ""  # 검색용 보조 필드 (텍스트만)
    
    # 출처 정보
    source_id: str = ""  # Textbook.id 또는 Exam.id
    source_type: SourceType = SourceType.TEXTBOOK  # Enum 사용
    
    # 태그 정보 (리스트)
    tags: List[Tag] = field(default_factory=list)
    
    # 메타데이터
    created_at: Optional[datetime] = None
    creator: str = ""
    
    # 원본 HWP 파일 정보 (참고용)
    original_hwp_path: Optional[str] = None  # 원본 파일 경로
    problem_index: int = 0  # 원본 파일 내 문제 순서 (1부터 시작)
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환 (MongoDB 저장용)"""
        return {
            '_id': self.id,
            'content_raw_file_id': self.content_raw_file_id,
            'content_text': self.content_text,
            'source_id': self.source_id,
            'source_type': self.source_type.value,  # Enum 값을 문자열로 저장
            'tags': [tag.to_dict() for tag in self.tags],
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'creator': self.creator,
            'original_hwp_path': self.original_hwp_path,
            'problem_index': self.problem_index
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Problem':
        """딕셔너리에서 생성"""
        # tags 리스트 복원
        tags = []
        if data.get('tags'):
            tags = [Tag.from_dict(tag_data) for tag_data in data['tags']]
        
        # source_type Enum 복원
        source_type = SourceType.TEXTBOOK
        if data.get('source_type'):
            try:
                source_type = SourceType(data['source_type'])
            except ValueError:
                source_type = SourceType.TEXTBOOK
        
        # created_at 복원
        created_at = None
        if data.get('created_at'):
            if isinstance(data['created_at'], str):
                created_at = datetime.fromisoformat(data['created_at'])
            else:
                created_at = data['created_at']
        
        return cls(
            id=data.get('_id'),
            content_raw_file_id=data.get('content_raw_file_id'),
            content_text=data.get('content_text', ''),
            source_id=data.get('source_id', ''),
            source_type=source_type,
            tags=tags,
            created_at=created_at,
            creator=data.get('creator', ''),
            original_hwp_path=data.get('original_hwp_path'),
            problem_index=data.get('problem_index', 0)
        )


@dataclass
class Textbook:
    """교재 메타데이터 모델"""
    id: Optional[str] = None  # MongoDB ObjectId
    name: str = ""  # 교재명 (예: "쎈수학")
    subject: str = ""  # 과목 (예: "공통수학1")
    major_unit: str = ""  # 대단원 (예: "다항식")
    sub_unit: Optional[str] = None  # 소단원 (예: "인수분해")
    created_at: Optional[datetime] = None
    parsed_at: Optional[datetime] = None  # HWP 파싱 완료 시각
    is_parsed: bool = False  # 파싱 완료 여부
    problem_count: int = 0  # 연결된 Problem 개수
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환"""
        return {
            '_id': self.id,
            'name': self.name,
            'subject': self.subject,
            'major_unit': self.major_unit,
            'sub_unit': self.sub_unit,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'parsed_at': self.parsed_at.isoformat() if self.parsed_at else None,
            'is_parsed': self.is_parsed,
            'problem_count': self.problem_count
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Textbook':
        """딕셔너리에서 생성"""
        created_at = None
        if data.get('created_at'):
            if isinstance(data['created_at'], str):
                created_at = datetime.fromisoformat(data['created_at'])
            else:
                created_at = data['created_at']
        
        parsed_at = None
        if data.get('parsed_at'):
            if isinstance(data['parsed_at'], str):
                parsed_at = datetime.fromisoformat(data['parsed_at'])
            else:
                parsed_at = data['parsed_at']
        
        return cls(
            id=data.get('_id'),
            name=data.get('name', ''),
            subject=data.get('subject', ''),
            major_unit=data.get('major_unit', ''),
            sub_unit=data.get('sub_unit'),
            created_at=created_at,
            parsed_at=parsed_at,
            is_parsed=data.get('is_parsed', False),
            problem_count=data.get('problem_count', 0)
        )


@dataclass
class Exam:
    """기출 메타데이터 모델"""
    id: Optional[str] = None  # MongoDB ObjectId
    grade: str = ""  # 학년 (예: "1학년")
    semester: str = ""  # 학기 (예: "1학기")
    exam_type: str = ""  # 유형 (예: "중간고사")
    school_name: str = ""  # 학교명
    year: str = ""  # 연도 (예: "2025")
    created_at: Optional[datetime] = None
    parsed_at: Optional[datetime] = None  # HWP 파싱 완료 시각
    is_parsed: bool = False  # 파싱 완료 여부
    problem_count: int = 0  # 연결된 Problem 개수
    
    def to_dict(self) -> dict:
        """딕셔너리로 변환"""
        return {
            '_id': self.id,
            'grade': self.grade,
            'semester': self.semester,
            'exam_type': self.exam_type,
            'school_name': self.school_name,
            'year': self.year,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'parsed_at': self.parsed_at.isoformat() if self.parsed_at else None,
            'is_parsed': self.is_parsed,
            'problem_count': self.problem_count
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Exam':
        """딕셔너리에서 생성"""
        created_at = None
        if data.get('created_at'):
            if isinstance(data['created_at'], str):
                created_at = datetime.fromisoformat(data['created_at'])
            else:
                created_at = data['created_at']
        
        parsed_at = None
        if data.get('parsed_at'):
            if isinstance(data['parsed_at'], str):
                parsed_at = datetime.fromisoformat(data['parsed_at'])
            else:
                parsed_at = data['parsed_at']
        
        return cls(
            id=data.get('_id'),
            grade=data.get('grade', ''),
            semester=data.get('semester', ''),
            exam_type=data.get('exam_type', ''),
            school_name=data.get('school_name', ''),
            year=data.get('year', ''),
            created_at=created_at,
            parsed_at=parsed_at,
            is_parsed=data.get('is_parsed', False),
            problem_count=data.get('problem_count', 0)
        )


@dataclass
class Worksheet:
    """
    학습지 메타데이터 모델

    - HWP/PDF 파일은 GridFS에 저장하고, 여기에는 file_id만 저장합니다.
    - problem_ids / numbered는 "학습지 구성"을 재현/추적하기 위한 최소 정보입니다.
    """

    id: Optional[str] = None  # MongoDB ObjectId

    title: str = ""  # 학습지명
    grade: str = ""  # 예: 초5, 중1, 고2
    type_text: str = ""  # 예: 내신기출 / 시중교재 / 통합
    creator: str = ""  # 출제자명(표시용)

    created_at: Optional[datetime] = None

    # 구성 정보
    problem_ids: List[str] = field(default_factory=list)  # 문제 id 리스트(순서 포함)
    numbered: List[dict] = field(default_factory=list)  # [{'no': 1, 'problem_id': '...'}, ...]

    # 파일 참조 (GridFS)
    hwp_file_id: Optional[str] = None
    pdf_file_id: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "_id": self.id,
            "title": self.title,
            "grade": self.grade,
            "type_text": self.type_text,
            "creator": self.creator,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "problem_ids": list(self.problem_ids or []),
            "numbered": list(self.numbered or []),
            "hwp_file_id": self.hwp_file_id,
            "pdf_file_id": self.pdf_file_id,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Worksheet":
        created_at = None
        if data.get("created_at"):
            if isinstance(data["created_at"], str):
                created_at = datetime.fromisoformat(data["created_at"])
            else:
                created_at = data["created_at"]

        return cls(
            id=data.get("_id"),
            title=data.get("title", "") or "",
            grade=data.get("grade", "") or "",
            type_text=data.get("type_text", "") or "",
            creator=data.get("creator", "") or "",
            created_at=created_at,
            problem_ids=list(data.get("problem_ids") or []),
            numbered=list(data.get("numbered") or []),
            hwp_file_id=data.get("hwp_file_id"),
            pdf_file_id=data.get("pdf_file_id"),
        )


@dataclass
class Student:
    """
    학생 모델

    UI/업무 필드(요구사항):
    - 학년, 상태(재원/휴원/퇴원), 학생이름, 학교명, 학부모 연락처, 학생 연락처

    내부 필드:
    - created_at/updated_at/deleted_at(소프트 삭제)
    """

    id: Optional[str] = None  # MongoDB ObjectId

    grade: str = ""  # 예: 초4, 중3, 고2
    status: str = "재원"  # 재원 | 휴원 | 퇴원
    name: str = ""
    school_name: str = ""
    parent_phone: str = ""
    student_phone: str = ""

    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    deleted_at: Optional[datetime] = None

    def to_dict(self) -> dict:
        return {
            "_id": self.id,
            "grade": self.grade,
            "status": self.status,
            "name": self.name,
            "school_name": self.school_name,
            "parent_phone": self.parent_phone,
            "student_phone": self.student_phone,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
            "deleted_at": self.deleted_at.isoformat() if self.deleted_at else None,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "Student":
        def _dt(v):
            if not v:
                return None
            if isinstance(v, str):
                try:
                    return datetime.fromisoformat(v)
                except Exception:
                    return None
            return v

        return cls(
            id=data.get("_id"),
            grade=data.get("grade", "") or "",
            status=data.get("status", "") or "",
            name=data.get("name", "") or "",
            school_name=data.get("school_name", "") or "",
            parent_phone=data.get("parent_phone", "") or "",
            student_phone=data.get("student_phone", "") or "",
            created_at=_dt(data.get("created_at")),
            updated_at=_dt(data.get("updated_at")),
            deleted_at=_dt(data.get("deleted_at")),
        )


@dataclass
class SchoolClass:
    """
    반(클래스) 모델

    UI/업무 필드:
    - 학년, 반명, 담당강사, 비고, 소속 학생 ID 목록

    내부 필드:
    - created_at/updated_at/deleted_at(소프트 삭제)
    """

    id: Optional[str] = None  # MongoDB ObjectId

    grade: str = ""  # 예: 초4, 중1, 고2
    name: str = ""  # 반명 (예: 1반, A반)
    teacher: str = ""  # 담당강사
    note: str = ""
    student_ids: List[str] = field(default_factory=list)  # 소속 학생 ObjectId 문자열 목록

    created_at: Optional[datetime] = None
    updated_at: Optional[datetime] = None
    deleted_at: Optional[datetime] = None

    def to_dict(self) -> dict:
        return {
            "_id": self.id,
            "grade": self.grade,
            "name": self.name,
            "teacher": self.teacher,
            "note": self.note,
            "student_ids": list(self.student_ids or []),
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at.isoformat() if self.updated_at else None,
            "deleted_at": self.deleted_at.isoformat() if self.deleted_at else None,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "SchoolClass":
        def _dt(v):
            if not v:
                return None
            if isinstance(v, str):
                try:
                    return datetime.fromisoformat(v)
                except Exception:
                    return None
            return v

        ids_raw = data.get("student_ids") or []
        student_ids = [str(x) for x in ids_raw if x]

        return cls(
            id=data.get("_id"),
            grade=(data.get("grade") or "").strip(),
            name=(data.get("name") or "").strip(),
            teacher=(data.get("teacher") or "").strip(),
            note=(data.get("note") or "").strip(),
            student_ids=student_ids,
            created_at=_dt(data.get("created_at")),
            updated_at=_dt(data.get("updated_at")),
            deleted_at=_dt(data.get("deleted_at")),
        )


@dataclass
class SavedReport:
    """
    저장된 학습 보고서(학생별, 기간별 스냅샷).

    - student_id: 대상 학생
    - period_start / period_end: 보고 기간 (YYYY-MM-DD)
    - comment: 학습코멘트
    - snapshot: 집계 결과 {total_worksheets, total_questions, total_correct, average_rate_pct, unit_stats: [...]}
    """

    id: Optional[str] = None
    student_id: str = ""
    period_start: str = ""  # YYYY-MM-DD
    period_end: str = ""   # YYYY-MM-DD
    comment: str = ""
    created_at: Optional[datetime] = None
    snapshot: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict:
        return {
            "_id": self.id,
            "student_id": (self.student_id or "").strip(),
            "period_start": (self.period_start or "").strip(),
            "period_end": (self.period_end or "").strip(),
            "comment": (self.comment or "").strip(),
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "snapshot": dict(self.snapshot or {}),
        }

    @classmethod
    def from_dict(cls, data: dict) -> "SavedReport":
        created_at = None
        if data.get("created_at"):
            v = data["created_at"]
            if isinstance(v, str):
                try:
                    created_at = datetime.fromisoformat(v)
                except Exception:
                    pass
            else:
                created_at = v
        return cls(
            id=data.get("_id"),
            student_id=(data.get("student_id") or "").strip(),
            period_start=(data.get("period_start") or "").strip(),
            period_end=(data.get("period_end") or "").strip(),
            comment=(data.get("comment") or "").strip(),
            created_at=created_at,
            snapshot=dict(data.get("snapshot") or {}),
        )

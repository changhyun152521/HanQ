"""
Problem 조회 및 관리 Service

Problem 데이터 조회, 검증, 요약 정보 제공을 담당합니다.
- Textbook/Exam ID 기준 Problem 목록 조회
- 파싱 품질 검증용 요약 정보 제공
- UI 연동을 고려한 데이터 구조 반환
"""
from typing import List, Optional, Dict, Any, Callable
import re
import os
import tempfile
from database.repositories import ProblemRepository, TextbookRepository, ExamRepository
from database.sqlite_connection import SQLiteConnection
from core.models import Problem, SourceType, Tag


class ProblemService:
    """Problem 조회 및 관리 서비스"""

    def _extract_best_preview_text(self, raw_text: str) -> str:
        """
        조회 단계에서만 사용하는 "미리보기용 텍스트" 추출.

        목적:
        - 파싱 과정(HWP 선택/저장/DB 저장)은 그대로 둠
        - content_text 안에 해설/정답이 섞이거나 순서가 뒤집혀도(주석 story 영향 등)
          "문제처럼 보이는" 구간을 우선적으로 프리뷰에 사용
        """
        if not raw_text:
            return ""

        lines = raw_text.splitlines()

        # 후보 구간들을 만들어 점수로 선택
        candidates = [raw_text]

        # 1) 구분선(긴 선) 기준 분리 후보
        sep_re = re.compile(r"^\s*[-─━―—=_]{8,}\s*$")
        for i, line in enumerate(lines):
            if sep_re.match(line):
                before = "\n".join(lines[:i]).strip()
                after = "\n".join(lines[i + 1 :]).strip()
                if before:
                    candidates.append(before)
                if after:
                    candidates.append(after)
                break

        # 2) 해설 시작(정답/풀이/해설) 기준 분리 후보
        sol_re = re.compile(r"^\s*(?:\d+\s*[\.\)]\s*)?(정답|풀이|해설)\b")
        for i, line in enumerate(lines):
            if sol_re.match(line):
                before = "\n".join(lines[:i]).strip()
                after = "\n".join(lines[i:]).strip()
                if before:
                    candidates.append(before)
                if after:
                    candidates.append(after)
                break

        def score(s: str) -> int:
            s2 = (s or "").strip()
            if not s2:
                return -10_000

            # 문제/해설 판별용 간단 휴리스틱
            good = 0
            bad = 0

            # 문제 문장에 자주 등장
            if re.search(r"(구하|계산|다음|값|최댓값|최솟값|증명|그래프|방정식|부등식)", s2):
                good += 3
            if "?" in s2:
                good += 2
            if re.search(r"[xy]\s*=", s2):
                good += 1

            # 해설에 자주 등장
            if re.search(r"(정답|풀이|해설|따라서|그러므로|이므로|결론)", s2):
                bad += 4

            # 너무 짧으면 프리뷰로 의미가 없으니 감점
            if len(s2) < 20:
                bad += 5

            return good - bad

        best = max(candidates, key=score)
        return best.strip()
    
    def __init__(self, db_connection: SQLiteConnection):
        """
        ProblemService 초기화
        
        Args:
            db_connection: DB 연결 인스턴스
        """
        self.db_connection = db_connection
        self.problem_repo = ProblemRepository(db_connection)
        self.textbook_repo = TextbookRepository(db_connection)
        self.exam_repo = ExamRepository(db_connection)
    
    def get_problems_by_source(
        self,
        source_id: str,
        source_type: SourceType
    ) -> List[Dict[str, Any]]:
        """
        Textbook 또는 Exam ID 기준으로 Problem 목록 조회
        
        Args:
            source_id: Textbook.id 또는 Exam.id
            source_type: SourceType Enum
            
        Returns:
            Problem 정보 리스트 (UI 연동용 구조)
            [
                {
                    'problem_id': str,
                    'problem_index': int,
                    'content_text_preview': str,  # 앞 200자
                    'has_content_raw': bool,
                    'tags': List[Dict] or None,
                    'created_at': str or None,
                    'original_hwp_path': str or None
                },
                ...
            ]
            
        Raises:
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 조회할 수 없습니다."
            )
        
        # Problem 목록 조회
        problems = self.problem_repo.find_by_source(source_id, source_type)
        
        # UI 연동용 구조로 변환
        result: List[Dict[str, Any]] = []
        for problem in problems:
            # content_text 미리보기 (앞 200자)
            # ✅ 조회 단계에서만 프리뷰 텍스트를 "문제처럼 보이는 구간"으로 보정
            raw_text_for_preview = self._extract_best_preview_text(problem.content_text or "")

            # - 목록에서 개행/공백만 있으면 "빈칸처럼" 보일 수 있어 공백을 압축한 값을 사용
            compact_text = " ".join(raw_text_for_preview.split())  # 모든 공백/개행을 단일 공백으로 정규화

            if compact_text:
                preview = compact_text[:200]
                if len(compact_text) > 200:
                    preview += "..."
            else:
                # 텍스트가 없더라도 "미리보기 없음"을 표시해 UI가 비어 보이지 않게 함
                if problem.content_raw_file_id is not None:
                    preview = "(텍스트 미리보기 없음: 수식/그림/표 위주)"
                else:
                    preview = "(텍스트 미리보기 없음)"
            
            # tags 정보 변환
            tags_data = None
            if problem.tags:
                tags_data = [tag.to_dict() for tag in problem.tags]

            # difficulty 추출 (UI 편의용): 대표 태그 우선, 없으면 fallback
            difficulty = None
            if problem.tags and len(problem.tags) > 0:
                difficulty = problem.tags[0].difficulty or None
                if not difficulty:
                    for t in problem.tags:
                        if t.difficulty:
                            difficulty = t.difficulty
                            break

            # 단원/과목 추출 (UI 편의용)
            subject = None
            major_unit = None
            sub_unit = None
            if problem.tags and len(problem.tags) > 0:
                primary = problem.tags[0]
                subject = primary.subject or None
                major_unit = getattr(primary, "major_unit", None) or None
                sub_unit = getattr(primary, "sub_unit", None) or None
                # 레거시 unit만 있는 경우도 표시할 수 있게 보조로 유지
                if (not major_unit) and getattr(primary, "unit", None):
                    major_unit = primary.unit

            unit_display = ""
            if major_unit and sub_unit:
                unit_display = f"{major_unit} > {sub_unit}"
            elif major_unit:
                unit_display = str(major_unit)
            
            result.append({
                'problem_id': problem.id,
                'problem_index': problem.problem_index,
                'content_text_preview': preview,
                'has_content_raw': problem.content_raw_file_id is not None,
                'tags': tags_data,
                'difficulty': difficulty,
                'subject': subject,
                'major_unit': major_unit,
                'sub_unit': sub_unit,
                'unit_display': unit_display,
                'created_at': problem.created_at.isoformat() if problem.created_at else None,
                'original_hwp_path': problem.original_hwp_path
            })
        
        # problem_index 기준 정렬
        result.sort(key=lambda x: x['problem_index'])
        
        return result

    def generate_previews_for_source(
        self,
        source_id: str,
        source_type: SourceType,
        only_missing: bool = True,
        progress_callback: Optional[Callable[[int, int], bool]] = None,
    ) -> Dict[str, int]:
        """
        원본 HWP 블록(GridFS)으로부터 content_text(미리보기/검색용)를 일괄 생성합니다.

        설계 원칙:
        - DB 생성(파싱) 속도 최우선 → 파싱 중 텍스트 추출을 생략할 수 있음
        - 이 함수는 사용자가 원할 때만 실행 (일괄 생성)
        - 파싱(선택 기반 블록 저장)에는 개입하지 않음

        Args:
            source_id: Textbook.id 또는 Exam.id
            source_type: SourceType
            only_missing: True면 content_text가 비어있는 문제만 생성
            progress_callback: (current, total) -> 계속 여부(bool). None이면 항상 계속.

        Returns:
            {'total': int, 'updated': int, 'skipped': int, 'failed': int}
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 수정할 수 없습니다."
            )

        problems = self.problem_repo.find_by_source(source_id, source_type)
        targets: List[Problem] = []
        skipped = 0
        for p in problems:
            if not p.content_raw_file_id:
                skipped += 1
                continue
            if only_missing and p.content_text and p.content_text.strip():
                skipped += 1
                continue
            targets.append(p)

        total = len(targets)
        updated = 0
        failed = 0

        # lazy import: HWP 환경이 없으면 기능만 실패하고 앱은 유지
        from processors.hwp.hwp_reader import HWPReader

        with HWPReader() as hwp:
            for idx, p in enumerate(targets, start=1):
                if progress_callback is not None:
                    cont = progress_callback(idx, total)
                    if cont is False:
                        break

                temp_path = None
                try:
                    hwp_bytes = self.problem_repo.get_content_raw(p.id) if p.id else None
                    if not hwp_bytes:
                        failed += 1
                        continue

                    fd, temp_path = tempfile.mkstemp(prefix=f"preview_{p.id}_", suffix=".hwp")
                    os.close(fd)
                    with open(temp_path, "wb") as f:
                        f.write(hwp_bytes)

                    if not hwp.open_document(temp_path):
                        failed += 1
                        continue

                    text = hwp.get_text_from_document()
                    hwp.close_document()

                    # 텍스트를 못 뽑았으면 실패로 집계 (update 여부와 무관)
                    if not text or not text.strip():
                        failed += 1
                        continue

                    p.content_text = text
                    # DB 업데이트
                    if not self.problem_repo.update(p):
                        # update()가 modified_count 기반이라 드물게 False가 나올 수 있음.
                        # 여기서는 텍스트가 추출되었고 예외가 없었다면 성공으로 간주합니다.
                        updated += 1
                    else:
                        updated += 1
                except Exception:
                    failed += 1
                finally:
                    try:
                        if temp_path and os.path.exists(temp_path):
                            os.remove(temp_path)
                    except Exception:
                        pass

        return {"total": total, "updated": updated, "skipped": skipped, "failed": failed}

    def _get_or_create_primary_tag(self, problem: Problem) -> Tag:
        """
        문제당 대표 태그(0번)를 보장합니다.
        - 기출DB는 문제별 단원/난이도 태깅이 필요하므로, Tag를 1개로 모아가는 편이 운영이 단순합니다.
        """
        if not problem.tags or len(problem.tags) == 0:
            problem.tags = [Tag(subject="", grade="")]
            return problem.tags[0]

        # "대표 태그"는 0번으로 정규화 (이미 값이 들어있는 태그가 있으면 그걸 앞으로)
        def has_any_value(t: Tag) -> bool:
            return bool(
                (t.subject or "").strip()
                or (t.grade or "").strip()
                or getattr(t, "major_unit", None)
                or getattr(t, "sub_unit", None)
                or getattr(t, "unit", None)
                or getattr(t, "difficulty", None)
            )

        for i, t in enumerate(problem.tags):
            if has_any_value(t):
                if i != 0:
                    problem.tags[0], problem.tags[i] = problem.tags[i], problem.tags[0]
                break

        return problem.tags[0]

    def set_problem_difficulty(self, problem_id: str, difficulty: Optional[str]) -> bool:
        """
        Problem 1건의 난이도(킬/상/중/하)를 저장합니다.

        구현 원칙:
        - 현 모델(`Tag.difficulty`)을 그대로 사용
        - 다른 태그 구조는 건드리지 않음
        - 난이도는 "첫 번째로 발견되는 difficulty"를 갱신(없으면 새 Tag 1개 추가)

        Args:
            problem_id: Problem ID
            difficulty: '킬' | '상' | '중' | '하' | None (미지정)

        Returns:
            성공 여부
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 수정할 수 없습니다."
            )

        problem = self.problem_repo.find_by_id(problem_id)
        if not problem:
            return False

        tag = self._get_or_create_primary_tag(problem)
        tag.difficulty = difficulty or None

        # 정규화: difficulty는 대표 태그 1곳에만 유지
        if problem.tags and len(problem.tags) > 1:
            for t in problem.tags[1:]:
                t.difficulty = None
        return self.problem_repo.update(problem)

    def set_problem_unit(
        self,
        problem_id: str,
        subject: str,
        major_unit: str,
        sub_unit: Optional[str] = None,
        grade: Optional[str] = None
    ) -> bool:
        """
        Problem 1건의 과목/단원(대단원/소단원)을 저장합니다.

        목표(빠른 실무 태깅):
        - 기출 HWP는 단원이 섞일 수 있어 Problem별로 단원을 저장해야 함
        - 대표 Tag(0번)에 subject/grade/major/sub/difficulty를 모아 운영 단순화
        - 레거시 호환을 위해 unit 문자열도 같이 채움
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 수정할 수 없습니다."
            )

        problem = self.problem_repo.find_by_id(problem_id)
        if not problem:
            return False

        tag = self._get_or_create_primary_tag(problem)
        tag.subject = subject or ""
        tag.grade = grade or tag.grade or ""

        major = (major_unit or "").strip()
        sub = (sub_unit or "").strip()
        tag.major_unit = major if major else None
        tag.sub_unit = sub if sub else None

        # 레거시 unit도 같이 채워서, 기존 UI/문서 호환
        if tag.major_unit and tag.sub_unit:
            tag.unit = f"{tag.major_unit} > {tag.sub_unit}"
        elif tag.major_unit:
            tag.unit = tag.major_unit
        else:
            tag.unit = None

        return self.problem_repo.update(problem)
    
    def get_parsing_summary(
        self,
        source_id: str,
        source_type: SourceType
    ) -> Dict[str, Any]:
        """
        파싱 품질 검증용 요약 정보 제공
        
        Args:
            source_id: Textbook.id 또는 Exam.id
            source_type: SourceType Enum
            
        Returns:
            {
                'total_problems': int,
                'empty_content_text_count': int,
                'missing_content_raw_count': int,
                'has_tags_count': int,
                'parsing_status': str,  # 'complete', 'partial', 'failed'
                'source_info': Dict  # Textbook 또는 Exam 정보
            }
            
        Raises:
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 조회할 수 없습니다."
            )
        
        # Problem 목록 조회
        problems = self.problem_repo.find_by_source(source_id, source_type)
        
        total = len(problems)
        empty_text_count = sum(1 for p in problems if not p.content_text or p.content_text.strip() == "")
        missing_raw_count = sum(1 for p in problems if not p.content_raw_file_id)
        has_tags_count = sum(1 for p in problems if p.tags and len(p.tags) > 0)
        
        # 파싱 상태 판단
        if total == 0:
            parsing_status = 'failed'
        elif empty_text_count == 0 and missing_raw_count == 0:
            parsing_status = 'complete'
        elif empty_text_count < total * 0.5 and missing_raw_count < total * 0.5:
            parsing_status = 'partial'
        else:
            parsing_status = 'failed'
        
        # 출처 정보 조회
        source_info = None
        if source_type == SourceType.TEXTBOOK:
            textbook = self.textbook_repo.find_by_id(source_id)
            if textbook:
                source_info = {
                    'name': textbook.name,
                    'subject': textbook.subject,
                    'major_unit': textbook.major_unit,
                    'sub_unit': textbook.sub_unit,
                    'is_parsed': textbook.is_parsed,
                    'problem_count': textbook.problem_count
                }
        elif source_type == SourceType.EXAM:
            exam = self.exam_repo.find_by_id(source_id)
            if exam:
                source_info = {
                    'grade': exam.grade,
                    'semester': exam.semester,
                    'exam_type': exam.exam_type,
                    'school_name': exam.school_name,
                    'year': exam.year,
                    'is_parsed': exam.is_parsed,
                    'problem_count': exam.problem_count
                }
        
        return {
            'total_problems': total,
            'empty_content_text_count': empty_text_count,
            'missing_content_raw_count': missing_raw_count,
            'has_tags_count': has_tags_count,
            'parsing_status': parsing_status,
            'source_info': source_info
        }
    
    def get_problem_detail(self, problem_id: str) -> Optional[Dict[str, Any]]:
        """
        Problem 상세 정보 조회
        
        Args:
            problem_id: Problem ID
            
        Returns:
            {
                'problem_id': str,
                'problem_index': int,
                'content_text': str,
                'has_content_raw': bool,
                'tags': List[Dict] or None,
                'source_id': str,
                'source_type': str,
                'created_at': str or None,
                'original_hwp_path': str or None,
                'creator': str
            } or None
            
        Raises:
            ConnectionError: MongoDB에 연결되지 않은 경우
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 조회할 수 없습니다."
            )
        
        problem = self.problem_repo.find_by_id(problem_id)
        if not problem:
            return None
        
        tags_data = None
        if problem.tags:
            tags_data = [tag.to_dict() for tag in problem.tags]
        
        return {
            'problem_id': problem.id,
            'problem_index': problem.problem_index,
            'content_text': problem.content_text,
            'has_content_raw': problem.content_raw_file_id is not None,
            'tags': tags_data,
            'source_id': problem.source_id,
            'source_type': problem.source_type.value,
            'created_at': problem.created_at.isoformat() if problem.created_at else None,
            'original_hwp_path': problem.original_hwp_path,
            'creator': problem.creator
        }

    def delete_problems_by_source(self, source_id: str, source_type: SourceType) -> int:
        """
        특정 출처(Textbook/Exam)에 연결된 Problem들을 모두 삭제합니다.

        - GridFS 원본(HWP)도 ProblemRepository.delete()에서 함께 삭제됩니다.
        - 출처 자체(Textbook/Exam) 삭제는 여기서 하지 않습니다.

        Returns:
            삭제된 Problem 개수
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 삭제할 수 없습니다."
            )

        deleted = 0
        problems = self.problem_repo.find_by_source(source_id, source_type)
        for p in problems:
            if p.id and self.problem_repo.delete(p.id):
                deleted += 1
        return deleted

    def delete_problems_by_ids(self, problem_ids: List[str]) -> int:
        """
        Problem ID 리스트를 받아 일괄 삭제합니다.

        Returns:
            삭제된 Problem 개수
        """
        if not self.db_connection.is_connected():
            raise ConnectionError(
                "DB에 연결되지 않았습니다. 데이터를 삭제할 수 없습니다."
            )

        deleted = 0
        for pid in problem_ids:
            if pid and self.problem_repo.delete(pid):
                deleted += 1
        return deleted

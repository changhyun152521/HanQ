"""
목업 데이터

개발 및 테스트를 위한 샘플 데이터를 제공합니다.
- 샘플 문제 데이터
- 샘플 태그 데이터
- 테스트용 HWP 파일 경로
"""
from datetime import datetime, timedelta
from core.models import Problem, Tag


def get_mock_problems() -> list[Problem]:
    """목업 문제 데이터 반환"""
    problems = []
    
    # 샘플 문제 1
    problems.append(Problem(
        id="1",
        content="다음 다항식의 곱셈을 계산하시오.\n\n(2x + 3)(x - 1) = ?",
        tag=Tag(
            subject="수학",
            grade="중1",
            unit="다항식의 연산",
            difficulty="중",
            source="쎈수학"
        ),
        created_at=datetime.now() - timedelta(days=2),
        creator="이창현T",
        worksheet_name="다항식의 곱셈 연습문제"
    ))
    
    # 샘플 문제 2
    problems.append(Problem(
        id="2",
        content="다음 이차방정식을 풀어라.\n\nx² - 5x + 6 = 0",
        tag=Tag(
            subject="수학",
            grade="중3",
            unit="이차방정식",
            difficulty="상",
            source="고쟁이"
        ),
        created_at=datetime.now() - timedelta(days=1),
        creator="김가나T",
        worksheet_name="이차방정식 기본 문제"
    ))
    
    # 샘플 문제 3
    problems.append(Problem(
        id="3",
        content="함수 f(x) = 2x + 1에 대하여 f(3)의 값을 구하시오.",
        tag=Tag(
            subject="수학",
            grade="고1",
            unit="함수",
            difficulty="하",
            source="블랙라벨"
        ),
        created_at=datetime.now() - timedelta(days=3),
        creator="나다라T",
        worksheet_name="함수의 값 구하기"
    ))
    
    # 샘플 문제 4
    problems.append(Problem(
        id="4",
        content="다음 부등식을 풀어라.\n\n2x - 3 > 5",
        tag=Tag(
            subject="수학",
            grade="중2",
            unit="일차부등식",
            difficulty="중",
            source="TOT"
        ),
        created_at=datetime.now() - timedelta(days=5),
        creator="강강강T",
        worksheet_name="일차부등식 연습"
    ))
    
    # 샘플 문제 5
    problems.append(Problem(
        id="5",
        content="다음 연립방정식을 풀어라.\n\n2x + y = 7\nx - y = 2",
        tag=Tag(
            subject="수학",
            grade="중2",
            unit="연립일차방정식",
            difficulty="상",
            source="쎈수학"
        ),
        created_at=datetime.now() - timedelta(days=7),
        creator="이창현T",
        worksheet_name="연립방정식 실전 문제"
    ))
    
    # 샘플 문제 6
    problems.append(Problem(
        id="6",
        content="다음 이차함수의 그래프의 꼭짓점을 구하시오.\n\ny = x² - 4x + 3",
        tag=Tag(
            subject="수학",
            grade="고2",
            unit="이차함수",
            difficulty="킬",
            source="고쟁이"
        ),
        created_at=datetime.now() - timedelta(days=10),
        creator="김가나T",
        worksheet_name="이차함수 그래프"
    ))
    
    return problems

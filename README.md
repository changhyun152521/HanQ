# HanQ

한글문서 기반의 학생별 맞춤 학습지부터 오답노트 및 개별 보고서 생성까지.

---

# CH-LMS 문제은행 시스템

HWP 문서에서 문제를 추출하고 태그를 붙여 MongoDB에 저장하는 문제은행 시스템입니다.

## 기능

- HWP 파일 불러오기 (현재 목업 모드)
- 문제 목록 테이블 표시
- 문제 선택 및 태그 입력
- 배치 선택 및 일괄 태그 적용
- 검색 및 필터 기능

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
python main.py
```

## 프로젝트 구조

```text
CH_LMS/
├── main.py                 # 메인 진입점
├── ui/                     # UI 컴포넌트
│   ├── main_window.py     # 메인 윈도우
│   ├── sidebar.py         # 사이드바
│   ├── problem_view.py    # 문제 목록 뷰
│   └── tag_form.py        # 태그 입력 폼
├── core/                   # 핵심 모델
│   └── models.py         # 데이터 모델
├── data/                   # 데이터
│   └── mock_data.py      # 목업 데이터
└── config/                 # 설정 파일
    ├── config.json        # 전역 설정
    └── tag_schema.json    # 태그 스키마
```

## 현재 상태

- ✅ UI 뼈대 구현 완료
- ✅ 목업 데이터 연동 완료
- ✅ HWP 파싱 기능 (마커 기반 문제 블록 절단)
- ✅ MongoDB 연결 및 GridFS 원본 저장

## 문서

- `docs/hwp-parsing-guide.md`: HWP 문제 블록 파싱/디버깅 가이드
- `docs/adr/0001-hwp-problem-block-extraction.md`: 추출 방식 결정 기록(ADR)
- `docs/로그인_API_MongoDB_Heroku_설치방법.md`: 로그인 API(Heroku + MongoDB Atlas) 설치
- `docs/데스크톱_exe_빌드_및_배포.md`: **배포용 exe 빌드 절차, 계정 생성, 배포 체크리스트**

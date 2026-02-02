# HWP 템플릿 머리말(헤더) 표가 사라지거나 “빈 머리말”이 추가 생성되는 문제 정리

이 문서는 CH_LMS에서 HWP 템플릿 기반으로 문제지를 생성할 때, **머리말(양쪽) 표가 저장 후 사라지거나** 또는 **같은 종류의 “빈 머리말(양쪽)”이 하나 더 생성되는** 현상을 재현/진단/해결한 내용을 정리합니다.

---

## 현상 요약

- **머리말(양쪽)에 있는 표가 통째로 사라짐**
  - 실행 중에는 표가 보이는데, 저장 후 열린 결과 파일에는 표가 없음
- **기존 머리말(양쪽) 표는 남아있지만, “빈 머리말(양쪽)”이 하나 더 생김**
  - 구역(Section)이 늘어나지 않았는데도 동일한 “양쪽 머리말”이 2개가 되는 케이스

---

## 원인(이번 케이스에서 확정된 것)

### 1) `HeaderFooter` 액션이 “머리말 편집 진입”이 아니라 “새(빈) 머리말 컨트롤 생성”으로 동작하는 케이스가 존재

- 한글(HWP) COM에서 `HAction.Execute("HeaderFooter", ...)`가 환경/문서 상태에 따라
  - 기존 머리말을 편집 모드로 열기도 하지만
  - **동일 타입(양쪽) 머리말 컨트롤을 새로 생성(빈 머리말)**해버리는 케이스가 있음
- 이 경우, 저장/닫기(정리) 시점에 “새로 생성된 빈 머리말”이 활성으로 선택되면서
  결과 파일에서 **머리말 표가 사라진 것처럼** 보이거나, 실제로 “양쪽 머리말 2개”가 됨

### 2) 헤더 내부에서의 `AllReplace/ExecReplace`는 표 컨트롤/선택 상태를 깨고, 저장/닫기와 결합될 때 문제를 악화시킬 수 있음

- 특히 머리말 표 셀 내부 텍스트 치환을 `AllReplace`로 수행한 뒤
  - 커서/선택 상태가 표 컨트롤로 승격되거나
  - 헤더 정리(닫기/저장) 과정에서 컨트롤이 손상/삭제/대체되는 패턴이 나타남

### 3) `ensure_main_body_focus()`는 내부에서 `CloseEx/Close`를 호출할 수 있어 “헤더 정리” 트리거가 될 수 있음

`processors/hwp/hwp_reader.py`의 `ensure_main_body_focus()`는 본문 복귀를 위해 `CloseEx`(Shift+Esc 동작)을 호출합니다.
헤더 편집/정리와 결합되면 예기치 않은 부작용이 발생할 수 있습니다.

---

## 해결(이번 프로젝트에서 성공한 최종 방식)

### 핵심: **HeaderFooter(머리말 편집 진입)를 사용하지 않는다**

머리말의 `HDR_*` 토큰 치환은 다음 방식으로 전환했습니다.

- **문서 전체(기본 스코프)에서 `RepeatFind`로 토큰을 찾아**
- **선택된 텍스트만 `Delete`로 제거하고**
- **`InsertText`(또는 텍스트 Paste 폴백)로 값을 삽입**

이 방식은:

- **머리말 “빈 컨트롤 추가 생성”을 유발하는 `HeaderFooter` 호출이 없고**
- **AllReplace/ExecReplace 사용을 피하므로**
- 저장 후에도 머리말 표가 유지되는 결과를 얻었습니다.

구현 위치:

- `services/worksheet/hwp_composer.py`
  - `_select_token_default_scope()`
  - `_replace_first_by_repeat_find()`
  - `_fill_header_markers_via_macro()` (현재는 이름이 “macro”지만, 내부는 HeaderFooter를 쓰지 않도록 변경됨)

---

## 템플릿 작성 규칙(권장)

- **머리말 표 셀 내부에 텍스트로 토큰을 둔다**
  - `HDR_DATE`, `HDR_TITLE`, `HDR_TEACHER`, `HDR_SCOPE`
- 본문에는 문제 삽입 위치로 `PROBLEMS_HERE`를 **정확히 1개** 넣는다
- 토큰은 가능한 한 **고유 문자열**로(문서 내 중복이 없게) 유지한다  
  (현재 구현은 “첫 번째 발견 1회 치환”에 최적화되어 있음)

---

## 진단/실험용 환경변수(재발 시 빠르게 고립하기)

`services/worksheet/hwp_composer.py`에 진단 모드가 포함되어 있습니다.

- `CH_LMS_HWP_DIAG_MODE`
  - `normal`: 기본 전체 플로우
  - `header_only`: 헤더 치환만 수행 후 저장
  - `body_only`: 본문(PROLEMS_HERE 제거 + 1문항 Paste)만 수행 후 저장
  - `combined`: header_only + body_only
- `CH_LMS_HWP_DIAG_VISIBLE=1`: 한글 창 표시
- `CH_LMS_HWP_DIAG_SLEEP_SEC=1.5`: 단계별 대기(초)
- `CH_LMS_HWP_DIAG_HEADER_EXIT=0`: 헤더 종료(닫기) 비활성화 (실험용)
- `CH_LMS_HWP_DIAG_HEADER_REPLACE=cell|allreplace`
  - 기본 권장: `cell`
  - `allreplace`는 헤더 표/컨트롤 문제를 재발시킬 수 있어 주의

---

## 빠른 체크리스트(다음 템플릿 적용 전/후)

- **템플릿 확인**
  - 머리말(양쪽)에 표가 있고, 표 셀에 `HDR_*` 토큰이 텍스트로 있는가
  - `PROBLEMS_HERE`가 본문에 1개만 있는가
- **생성 테스트**
  - 먼저 `header_only`로 머리말만 저장 테스트
  - 머리말이 유지되면 `normal`로 전체 생성 테스트
- **문제 재발 시**
  - “양쪽 머리말이 2개로 늘어나는지”부터 확인  
    (늘어난다면 HeaderFooter 계열 호출이 개입되고 있는지 의심)

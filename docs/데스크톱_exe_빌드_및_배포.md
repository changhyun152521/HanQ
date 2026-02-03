# HanQ 데스크톱 exe 빌드 및 배포

배포용 exe는 **Heroku 로그인 API**에 연결되며, **사용자님이 생성한 계정으로만** 로그인할 수 있습니다.

---

## 1. 사전 준비

- **Python 3** (프로젝트와 동일한 버전)
- **가상환경** 권장: `python -m venv .venv` 후 활성화
- **의존성 설치**: `pip install -r requirements.txt`
- **PyInstaller 설치**: `pip install pyinstaller`

---

## 2. exe 빌드 절차

### 2-0. 한글 경로에서 빌드하기

프로젝트에 **한글 경로 대응**이 적용되어 있어, **한글이 포함된 경로**(예: `C:\Users\이창현\Desktop\CH_LMS`)에서도 빌드할 수 있습니다.

- 빌드 스크립트가 **Python UTF-8 모드**로 PyInstaller를 실행합니다.
- **커스텀 훅**(`pyinstaller_hooks/hook-PyQt5.QtWidgets.py`)이 현재 프로세스에서 PyQt5 플러그인 경로를 계산해 사용합니다.

그래도 **"Qt plugin directory does not exist"** 오류가 나면, 프로젝트를 한글이 없는 경로(예: `C:\CH_LMS`)로 복사한 뒤 그곳에서 빌드해 보세요.

### 2-1. 한 번만 확인할 것

- **config/config.deploy.json**의 `login_api.base_url`이 현재 Heroku 앱 주소와 같은지 확인  
  - 예: `https://hanq-ch-lms-99d87e53375f.herokuapp.com`  
  - Heroku 앱 이름을 바꾼 경우 여기를 수정한 뒤 빌드

### 2-2. 빌드 실행

**PowerShell**에서 프로젝트 루트(`CH_LMS`)로 이동한 뒤:

```powershell
.\scripts\build_exe.ps1
```

- 스크립트가 **배포용 설정**을 적용한 뒤 PyInstaller를 실행하고, 끝나면 **로컬 개발용 config.json**을 복원합니다.
- 빌드 결과: **`dist\HanQ\`** 폴더  
  - 실행 파일: `dist\HanQ\HanQ.exe`  
  - 같은 폴더의 `config`, `templates` 등이 함께 있어야 합니다.

### 2-3. 빌드 후 확인

1. **Qt 플러그인 폴더 확인**  
   `dist\HanQ\PyQt5\Qt5\plugins\platforms` 폴더가 있고, 그 안에 `qwindows.dll` 등이 있어야 exe가 실행됩니다. 없으면 "no Qt platform plugin could be initialized" 오류가 납니다.
2. `dist\HanQ\HanQ.exe` 더블클릭으로 실행
3. 로그인 화면에서 **Heroku에 올라간 계정**으로 로그인 시도
4. 로그인 성공 후 메인 화면까지 정상 동작하는지 확인

---

## 3. 배포 절차 (사용자에게 전달)

### 3-1. 배포 패키지 만들기

- **`dist\HanQ`** 폴더 **전체**를 ZIP 등으로 압축
- 이 폴더 안에 `HanQ.exe`, `config`, `templates`, 기타 dll/리소스가 모두 들어 있어야 함

### 3-2. 배포 채널

- **Google Drive / OneDrive / 내부 서버** 등에 압축 파일 업로드
- 또는 **GitHub Releases**에 올려서 다운로드 링크 제공

### 3-3. 사용자 안내

- 압축 해제 후 **HanQ.exe** 실행
- **로그인**: 사용자님이 전달한 **user_id / 비밀번호**로 로그인  
  - 계정이 없으면 로그인 불가 → 프로그램 이용 불가 (의도된 동작)

---

## 4. 계정 생성 (로그인 가능한 사용자 만들기)

로그인으로만 이용 가능하게 하려면, **이용할 사람마다 사용자님이 계정을 생성**해 주어야 합니다.

### 4-1. MongoDB Atlas에서 직접 추가

1. [MongoDB Atlas](https://cloud.mongodb.com) 로그인
2. **Database** → **Browse Collections**
3. 데이터베이스 **hanq** → 컬렉션 **users** 선택 (없으면 Create Database / Create Collection)
4. **Add Document** (Insert Document) 후 예시처럼 입력:

```json
{
  "user_id": "학생아이디",
  "password": "설정한비밀번호",
  "name": "이름",
  "created_at": "2025-02-03T00:00:00.000Z"
}
```

- `user_id`, `password`, `name`을 사용자에게 전달

### 4-2. HanQ 앱(관리자)으로 추가

- 이미 **admin** 계정으로 로그인한 상태에서  
  **관리 탭 → 회원 관리 → 회원 추가**로 새 계정 생성 후, 해당 **user_id / 비밀번호**를 사용자에게 전달

---

## 5. 배포 체크리스트

| 항목 | 확인 |
|------|------|
| Heroku 로그인 API 정상 동작 | 브라우저에서 `base_url` 접속 시 `{"ok":true,"service":"ch-lms-auth-api"}` 확인 |
| config.deploy.json의 base_url | 현재 Heroku 앱 URL과 일치 |
| 빌드 스크립트 실행 | `.\scripts\build_exe.ps1` 성공 |
| exe 단독 실행 | `dist\HanQ\HanQ.exe` 실행 후 로그인 화면 표시 |
| Heroku 계정으로 로그인 | 테스트 계정으로 로그인 성공 및 메인 화면 진입 |
| 배포 패키지 | `dist\HanQ` 폴더 전체 압축 후 전달 |
| 사용자 계정 | MongoDB(또는 관리자 기능)로 생성 후 user_id/비밀번호 전달 |

---

## 6. 로컬 개발과 배포 설정 분리

- **로컬 개발**: `config/config.json` → `base_url`을 `http://localhost:5000` 등으로 두고 사용
- **배포용 exe**: 빌드 시 `config.deploy.json`이 자동으로 적용되므로, exe는 항상 Heroku 주소로 로그인
- 빌드 스크립트가 끝나면 **config.json은 다시 로컬용으로 복원**되므로, 일상적인 개발에는 영향 없음

Heroku 앱 URL을 바꾼 경우 **config/config.deploy.json**만 수정한 뒤 다시 `.\scripts\build_exe.ps1` 실행하면 됩니다.

# HanQ 로그인 API — MongoDB Atlas + Heroku 설치 방법

HanQ 앱의 **로그인/회원 관리**는 **MongoDB Atlas**(유저 DB)와 **Heroku**(API 서버)로 동작합니다.  
아래 순서대로 진행하면 됩니다.

---

## 0. 로컬 MongoDB + Compass로 먼저 테스트하기 (선택)

Atlas/Heroku 배포 전에, 로컬 MongoDB와 **MongoDB Compass**로 API와 데이터를 확인하고 싶다면 이 단계를 먼저 진행하세요.

### 0-1. 로컬 MongoDB 실행

**방법 A — MongoDB Community Server 설치**

1. [MongoDB Community Server](https://www.mongodb.com/try/download/community) 에서 OS에 맞게 다운로드 후 설치.
2. 설치 시 **Install MongoDB as a Service** 옵션을 켜 두면 부팅 시 자동 실행됩니다.
3. 기본 포트 **27017** 로 실행됩니다.

**방법 B — Docker 사용**

```bash
docker run -d -p 27017:27017 --name mongo-local mongo:latest
```

### 0-2. MongoDB Compass 설치 (선택)

1. [MongoDB Compass](https://www.mongodb.com/try/download/compass) 다운로드 후 설치.
2. 실행 후 **Connect** 화면에서 연결 주소에 `mongodb://localhost:27017` 입력 후 **Connect**.
3. 나중에 API로 회원을 추가하면 **hanq** → **users** 컬렉션에서 문서를 확인·수정할 수 있습니다.

### 0-3. auth_api가 로컬 DB를 쓰도록 설정

1. **auth_api** 폴더에 `.env` 파일을 만듭니다 (없으면).
2. 아래 내용을 넣습니다 (로컬 MongoDB 사용):

```
# 로컬 MongoDB (Compass로 같은 DB 확인 가능)
MONGODB_URI=mongodb://localhost:27017
PORT=5000
```

3. 터미널에서 API 서버를 실행합니다:

```bash
cd auth_api
npm install
npm start
```

4. **config/config.json** 의 **login_api.base_url** 을 로컬 API로 설정합니다:

```json
"login_api": {
  "base_url": "http://localhost:5000"
}
```

### 0-4. 첫 회원 넣기 및 Compass로 확인

1. **첫 회원**은 로그인 전이므로 Compass에서 넣습니다.  
   Compass에서 **mongodb://localhost:27017** 접속 → **Create Database** → Database name `hanq`, Collection name `users` → **Create**.  
   **users** 컬렉션 → **Add Data** → **Insert Document** 에서 아래 JSON 넣고 Insert:

```json
{
  "user_id": "admin",
  "password": "test123",
  "name": "관리자",
  "created_at": "2025-01-01T00:00:00.000Z"
}
```

2. HanQ 앱을 실행한 뒤, 로그인 화면에서 `admin` / `test123` 으로 로그인합니다.  
   로그인 후 **관리 탭 → 회원 관리 → 회원 추가**로 다른 계정을 추가해 보면, Compass에서 **hanq** → **users** 를 새로고침해 추가된 문서를 확인할 수 있습니다.

3. 테스트가 끝나면 Atlas/Heroku 배포 시 **auth_api/.env** 의 **MONGODB_URI** 를 Atlas 연결 문자열로 바꾸고, **config.json** 의 **base_url** 을 Heroku 주소로 바꾸면 됩니다.

---

## 1. MongoDB Atlas 준비

### 1-1. 가입 및 클러스터 생성

1. [MongoDB Atlas](https://www.mongodb.com/cloud/atlas) 접속 후 **무료 가입**.
2. **Build a Database** → **M0 FREE** 선택 후 **Create**.
3. 리전은 가까운 곳(예: **Seoul** 또는 **Singapore**) 선택.
4. 클러스터 이름은 그대로 두거나 원하는 이름으로 변경 후 **Create**.

### 1-2. DB 사용자 생성

1. **Security** → **Database Access** → **Add New Database User**.
2. **Authentication Method**: Password.
3. **Username** / **Password** 지정 후 **Add User** (비밀번호는 꼭 저장해 두세요).

### 1-3. 네트워크 접근 허용

1. **Security** → **Network Access** → **Add IP Address**.
2. **Allow Access from Anywhere** (0.0.0.0/0) 선택 후 **Confirm** (Heroku에서 접속해야 하므로).

### 1-4. 연결 문자열 복사

1. **Database** → **Connect** → **Drivers** 선택.
2. **Connection string** 복사.  
   형식 예:  
   `mongodb+srv://<username>:<password>@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority`
3. `<username>` / `<password>`를 1-2에서 만든 DB 사용자 정보로 **실제 값으로 바꿔** 두세요.  
   (특수문자는 URL 인코딩 필요. 예: `@` → `%40`)

---

## 2. Heroku에 API 배포

### 2-1. Heroku 가입 및 CLI 설치

1. [Heroku](https://www.heroku.com) 가입.
2. [Heroku CLI](https://devcenter.heroku.com/articles/heroku-cli) 설치 후 터미널에서 `heroku login` 실행.

### 2-2. 앱 생성 및 배포

프로젝트 루트(`CH_LMS`)에서:

```bash
# Heroku 앱 생성 (이름은 원하는 대로, 예: ch-lms-auth)
heroku create ch-lms-auth

# MongoDB 연결 문자열 설정 (1-4에서 복사한 문자열 전체)
heroku config:set MONGODB_URI="mongodb+srv://사용자:비밀번호@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority"

# 배포 (Git 사용 시)
git add .
git commit -m "Add auth API for Heroku"
git push heroku main
```

- **main** 브랜치가 아니면 `git push heroku 실제브랜치이름:main` 으로 푸시.

### 2-3. 배포 확인

- 브라우저에서 `https://ch-lms-auth.herokuapp.com` (본인이 만든 앱 이름으로) 접속.
- `{"ok":true,"service":"ch-lms-auth-api"}` 가 보이면 정상입니다.

---

## 3. HanQ 앱 설정

1. **config/config.json** 을 엽니다.
2. **login_api.base_url** 을 Heroku 앱 주소로 바꿉니다.

```json
"login_api": {
  "base_url": "https://ch-lms-auth.herokuapp.com"
}
```

- `ch-lms-auth` 부분을 본인 Heroku 앱 이름으로 수정하세요.
- **끝에 슬래시(/) 붙이지 마세요.**

---

## 4. 첫 회원 추가 (MongoDB에 테스트 유저 넣기)

API가 배포된 상태에서:

1. HanQ 앱을 실행합니다.
2. 로그인 화면이 나오면, **회원이 한 명도 없으면** 먼저 MongoDB에 직접 한 명 넣거나,  
   **관리자용 회원 추가**가 있다면 그 기능으로 추가합니다.

**MongoDB Atlas에서 직접 넣는 방법:**

1. Atlas **Database** → **Browse Collections**.
2. **Create Database**: Database name `hanq`, Collection name `users`.
3. **users** 컬렉션 → **Add Document** (Insert Document).
4. 아래 JSON 한 번에 넣거나, 필드만 맞춰서 입력:

```json
{
  "user_id": "admin",
  "password": "원하는비밀번호",
  "name": "관리자",
  "created_at": "2025-01-01T00:00:00.000Z"
}
```

5. **Insert** 후 HanQ 로그인 화면에서 `admin` / 위에서 넣은 비밀번호로 로그인해 보세요.

---

## 5. 요약 체크리스트

- [ ] MongoDB Atlas 가입, M0 클러스터 생성
- [ ] DB 사용자 생성, 네트워크 0.0.0.0/0 허용
- [ ] 연결 문자열 복사 후 `<username>`, `<password>` 치환
- [ ] Heroku 가입, CLI 설치, `heroku login`
- [ ] `heroku create`, `heroku config:set MONGODB_URI=...`
- [ ] `git push heroku main` 으로 배포
- [ ] **config/config.json** 의 **login_api.base_url** 을 Heroku URL로 설정
- [ ] MongoDB `hanq.users` 에 테스트 계정 한 개 추가 후 HanQ 로그인 테스트

---

## 6. 로컬에서 Atlas만 연결해서 쓰기 (데이터를 Atlas에 배포)

이미 Atlas 클러스터가 있고, **로컬 PC에서 auth_api만 Atlas에 연결**해 로그인/회원 데이터를 Atlas에 두고 싶을 때:

1. **auth_api** 폴더에 `.env` 파일을 만들거나 엽니다.
2. 아래 한 줄을 넣습니다 (본인 Atlas 연결 문자열로 교체).  
   **끝에 `?retryWrites=true&w=majority` 가 없으면 붙여 주세요.**

```
MONGODB_URI=mongodb+srv://<사용자명>:<비밀번호>@<클러스터>.mongodb.net/?retryWrites=true&w=majority
PORT=5000
```

3. 터미널에서 auth_api 실행:

```bash
cd auth_api
npm install
npm start
```

4. **config/config.json** 의 **login_api.base_url** 을 `http://localhost:5000` 으로 두고 HanQ 앱에서 로그인 테스트.
5. Atlas **Database** → **Browse Collections** 에서 **hanq** → **users** 컬렉션이 생성되고, 회원 추가 시 문서가 쌓이는지 확인합니다.

- **주의:** `.env` 에 넣는 연결 문자열에는 비밀번호가 포함됩니다. `.env` 는 Git에 올리지 마세요. (이미 .gitignore 에 있음)  
- 연결 문자열을 다른 사람에게 보냈다면 Atlas **Database Access** 에서 해당 사용자 비밀번호를 **바로 변경**하세요.

---

## 7. 문제 해결

| 증상 | 확인할 것 |
|------|------------|
| "config.json에 로그인 API 주소를 설정해 주세요" | **config/config.json** 의 **login_api.base_url** 이 Heroku 주소인지, `your-app.herokuapp.com` 이 그대로 남아 있지 않은지 확인. |
| "DB 연결 실패" | Heroku **Config Vars** 에 **MONGODB_URI** 가 올바르게 설정되었는지, Atlas 비밀번호에 특수문자 있으면 URL 인코딩 했는지 확인. |
| 로그인/회원 목록 404 | Heroku 앱이 정상 배포되었는지, `https://앱이름.herokuapp.com/` 접속 시 `{"ok":true,...}` 가 나오는지 확인. |

추가로 Heroku 로그는 `heroku logs --tail` 로 확인할 수 있습니다.

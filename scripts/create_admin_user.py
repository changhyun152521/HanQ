"""
관리자 계정 생성 스크립트

MongoDB hanq.users 컬렉션에 관리자(admin) 계정을 한 번 넣습니다.
auth_api/.env 의 MONGODB_URI 를 사용하며, 없으면 mongodb://localhost:27017 을 씁니다.

실행: 프로젝트 루트(CH_LMS)에서
  python scripts/create_admin_user.py
"""
import os
import sys
from datetime import datetime, timezone

# 프로젝트 루트를 path에 추가
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

ADMIN_USER_ID = "admin"
ADMIN_PASSWORD = "admin123"
ADMIN_NAME = "관리자"
DB_NAME = "hanq"
COLLECTION_NAME = "users"


def _load_mongodb_uri() -> str:
    env_path = os.path.join(ROOT, "auth_api", ".env")
    if os.path.isfile(env_path):
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("MONGODB_URI=") and not line.startswith("#"):
                    value = line.split("=", 1)[1].strip().strip('"').strip("'")
                    if value:
                        return value
    return "mongodb://localhost:27017"


def main() -> None:
    try:
        from pymongo import MongoClient
    except ImportError:
        print("pymongo가 없습니다. pip install pymongo 후 다시 실행하세요.")
        sys.exit(1)

    uri = _load_mongodb_uri()
    print(f"MongoDB 연결: {uri[:50]}...")

    client = MongoClient(uri)
    db = client[DB_NAME]
    col = db[COLLECTION_NAME]

    if col.find_one({"user_id": ADMIN_USER_ID}):
        print(f"이미 '{ADMIN_USER_ID}' 계정이 있습니다. 비밀번호를 바꾸려면 MongoDB에서 직접 수정하세요.")
        return

    doc = {
        "user_id": ADMIN_USER_ID,
        "password": ADMIN_PASSWORD,
        "name": ADMIN_NAME,
        "created_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
    }
    col.insert_one(doc)
    print("관리자 계정이 생성되었습니다.")
    print(f"  아이디: {ADMIN_USER_ID}")
    print(f"  비밀번호: {ADMIN_PASSWORD}")
    print("  (보안을 위해 첫 로그인 후 비밀번호 변경을 권장합니다.)")


if __name__ == "__main__":
    main()

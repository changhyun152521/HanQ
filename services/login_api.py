"""
HanQ 로그인/회원 API 클라이언트 (Heroku 배포 API 호출)
"""
from __future__ import annotations

import json
import os
from typing import Any, Dict, List, Optional

# requests는 선택 의존성: 없으면 로그인/회원 API 호출 시 에러 메시지 반환
try:
    import requests
    _HAS_REQUESTS = True
except ImportError:
    _HAS_REQUESTS = False


def _load_config() -> Dict[str, Any]:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(root, "config", "config.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def get_base_url() -> str:
    cfg = _load_config()
    url = (cfg.get("login_api") or {}).get("base_url") or ""
    return url.rstrip("/")


def _parse_json_response(r: Any, request_url: str = "") -> tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    응답에서 JSON 파싱. (data, None) 또는 (None, error_message) 반환.
    빈 응답·비JSON 응답 시 에러 메시지 반환으로 'Expecting value: line 1 column 1' 방지.
    """
    url_hint = f"\n요청 주소: {request_url}" if request_url else ""
    if r.status_code != 200:
        if r.status_code == 404:
            return (
                None,
                f"서버 응답 오류 (HTTP 404). 요청한 주소에 API가 없습니다.{url_hint}\n\n"
                "→ 이 프로젝트의 auth_api(Node.js) 서버가 실행 중인지 확인하세요.\n"
                "  (auth_api 폴더에서 'npm start' 실행 후, config.json의 base_url이 해당 주소인지 확인)",
            )
        return None, f"서버 응답 오류 (HTTP {r.status_code}). 서버 주소와 서버 상태를 확인해 주세요.{url_hint}"
    text = (r.text or "").strip()
    if not text:
        return None, f"서버가 빈 응답을 반환했습니다. API 주소와 서버가 정상 동작하는지 확인해 주세요.{url_hint}"
    try:
        return json.loads(text), None
    except json.JSONDecodeError as e:
        return None, f"서버 응답이 JSON 형식이 아닙니다. (서버 주소·API 경로 확인) — {e}{url_hint}"


def _session_path() -> str:
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root, "config", ".hanq_session.json")


def load_session() -> Optional[Dict[str, Any]]:
    """저장된 로그인 세션 반환. 없거나 잘못되면 None."""
    path = _session_path()
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        uid = (data.get("user_id") or "").strip()
        if uid:
            return {"user_id": uid, "name": (data.get("name") or uid).strip() or uid}
    except Exception:
        pass
    return None


def save_session(user_id: str, name: str) -> None:
    path = _session_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"user_id": user_id, "name": name or user_id}, f, ensure_ascii=False)
    except Exception:
        pass


def clear_session() -> None:
    path = _session_path()
    try:
        if os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass


def login(user_id: str, password: str) -> Dict[str, Any]:
    """
    로그인 요청.
    반환: { "success": True/False, "name": "이름" 또는 "message": "에러 메시지" }
    """
    if not _HAS_REQUESTS:
        return {"success": False, "message": "requests 패키지가 없습니다. pip install requests 로 설치해 주세요."}
    base = get_base_url()
    if not base or "your-app.herokuapp.com" in base:
        return {"success": False, "message": "config.json에 로그인 API 주소(base_url)를 설정해 주세요."}
    url = f"{base}/login.php"
    try:
        r = requests.post(url, json={"user_id": user_id, "password": password}, timeout=30)
        r.encoding = "utf-8"
        data, err = _parse_json_response(r, url)
        if err:
            return {"success": False, "message": err}
    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"서버 연결 실패: {e}"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    if not data.get("success"):
        return {"success": False, "message": data.get("message", "로그인에 실패했습니다.")}
    return {"success": True, "name": data.get("name") or user_id, "user_id": data.get("user_id") or user_id}


def list_users() -> Dict[str, Any]:
    """
    회원 목록 조회.
    반환: { "success": True/False, "users": [ { "user_id", "name", "created_at" }, ... ] 또는 "message": "..." }
    """
    if not _HAS_REQUESTS:
        return {"success": False, "message": "requests 패키지가 없습니다.", "users": []}
    base = get_base_url()
    if not base or "your-app.herokuapp.com" in base:
        return {"success": False, "message": "config.json에 로그인 API 주소를 설정해 주세요.", "users": []}
    url = f"{base}/list_users.php"
    try:
        r = requests.get(url, timeout=30)
        r.encoding = "utf-8"
        data, err = _parse_json_response(r, url)
        if err:
            return {"success": False, "message": err, "users": []}
    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"서버 연결 실패: {e}", "users": []}
    except Exception as e:
        return {"success": False, "message": str(e), "users": []}
    if not data.get("success"):
        return {"success": False, "message": data.get("message", "목록 조회 실패"), "users": []}
    return {"success": True, "users": data.get("users") or []}


def add_user(user_id: str, password: str, name: str) -> Dict[str, Any]:
    """
    회원 추가.
    반환: { "success": True/False, "message": "..." }
    """
    if not _HAS_REQUESTS:
        return {"success": False, "message": "requests 패키지가 없습니다."}
    base = get_base_url()
    if not base or "your-app.herokuapp.com" in base:
        return {"success": False, "message": "config.json에 로그인 API 주소를 설정해 주세요."}
    url = f"{base}/add_user.php"
    try:
        r = requests.post(url, json={"user_id": user_id, "password": password, "name": name or user_id}, timeout=30)
        r.encoding = "utf-8"
        data, err = _parse_json_response(r, url)
        if err:
            return {"success": False, "message": err}
    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"서버 연결 실패: {e}"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    if not data.get("success"):
        return {"success": False, "message": data.get("message", "회원 추가에 실패했습니다.")}
    return {"success": True}


def update_user(
    current_user_id: str,
    password: str,
    new_user_id: Optional[str] = None,
    new_name: Optional[str] = None,
    new_password: Optional[str] = None,
) -> Dict[str, Any]:
    """
    회원 정보 수정 (이름, 아이디, 비밀번호).
    반환: { "success": True/False, "user_id": "...", "name": "..." 또는 "message": "..." }
    """
    if not _HAS_REQUESTS:
        return {"success": False, "message": "requests 패키지가 없습니다."}
    base = get_base_url()
    if not base or "your-app.herokuapp.com" in base:
        return {"success": False, "message": "config.json에 로그인 API 주소를 설정해 주세요."}
    url = f"{base}/update_user.php"
    payload = {"current_user_id": current_user_id, "password": password}
    if new_user_id is not None:
        payload["new_user_id"] = new_user_id
    if new_name is not None:
        payload["new_name"] = new_name
    if new_password is not None:
        payload["new_password"] = new_password
    try:
        r = requests.post(url, json=payload, timeout=30)
        r.encoding = "utf-8"
        data, err = _parse_json_response(r, url)
        if err:
            return {"success": False, "message": err}
    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"서버 연결 실패: {e}"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    if not data.get("success"):
        return {"success": False, "message": data.get("message", "정보 수정에 실패했습니다.")}
    return {
        "success": True,
        "user_id": data.get("user_id") or current_user_id,
        "name": data.get("name") or new_name or current_user_id,
    }


def admin_update_user(
    admin_user_id: str,
    admin_password: str,
    target_user_id: str,
    new_user_id: Optional[str] = None,
    new_name: Optional[str] = None,
    new_password: Optional[str] = None,
) -> Dict[str, Any]:
    """
    관리자용 회원 수정 (admin만 호출 가능).
    반환: { "success": True/False, "message": "..." }
    """
    if not _HAS_REQUESTS:
        return {"success": False, "message": "requests 패키지가 없습니다."}
    base = get_base_url()
    if not base or "your-app.herokuapp.com" in base:
        return {"success": False, "message": "config.json에 로그인 API 주소를 설정해 주세요."}
    url = f"{base}/admin_update_user.php"
    payload = {
        "admin_user_id": admin_user_id,
        "admin_password": admin_password,
        "target_user_id": target_user_id,
    }
    if new_user_id is not None:
        payload["new_user_id"] = new_user_id
    if new_name is not None:
        payload["new_name"] = new_name
    if new_password is not None:
        payload["new_password"] = new_password
    try:
        r = requests.post(url, json=payload, timeout=30)
        r.encoding = "utf-8"
        data, err = _parse_json_response(r, url)
        if err:
            return {"success": False, "message": err}
    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"서버 연결 실패: {e}"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    if not data.get("success"):
        return {"success": False, "message": data.get("message", "회원 수정에 실패했습니다.")}
    return {"success": True}


def delete_user(admin_user_id: str, admin_password: str, target_user_id: str) -> Dict[str, Any]:
    """
    관리자용 회원 삭제 (admin만 호출 가능).
    반환: { "success": True/False, "message": "..." }
    """
    if not _HAS_REQUESTS:
        return {"success": False, "message": "requests 패키지가 없습니다."}
    base = get_base_url()
    if not base or "your-app.herokuapp.com" in base:
        return {"success": False, "message": "config.json에 로그인 API 주소를 설정해 주세요."}
    url = f"{base}/delete_user.php"
    try:
        r = requests.post(
            url,
            json={
                "admin_user_id": admin_user_id,
                "admin_password": admin_password,
                "target_user_id": target_user_id,
            },
            timeout=30,
        )
        r.encoding = "utf-8"
        data, err = _parse_json_response(r, url)
        if err:
            return {"success": False, "message": err}
    except requests.exceptions.RequestException as e:
        return {"success": False, "message": f"서버 연결 실패: {e}"}
    except Exception as e:
        return {"success": False, "message": str(e)}
    if not data.get("success"):
        return {"success": False, "message": data.get("message", "회원 삭제에 실패했습니다.")}
    return {"success": True}

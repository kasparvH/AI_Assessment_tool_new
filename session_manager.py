import json
import os
import random
import string
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).parent
SESSIONS_DIR = ROOT / "sessions"
SESSIONS_DIR.mkdir(exist_ok=True)


def _session_path(code: str) -> Path:
    return SESSIONS_DIR / f"{code.upper()}.json"


def generate_session_code() -> str:
    prefix = "".join(random.choices(string.ascii_uppercase, k=3))
    suffix = "".join(random.choices(string.ascii_uppercase + string.digits, k=3))
    code = f"{prefix}-{suffix}"
    if _session_path(code).exists():
        return generate_session_code()
    return code


def save_session(session_code: str, state: dict) -> None:
    state = dict(state)
    state["last_saved_at"] = datetime.now().isoformat()
    _session_path(session_code).write_text(
        json.dumps(state, indent=2, default=str), encoding="utf-8"
    )


def load_session(session_code: str) -> dict | None:
    path = _session_path(session_code)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def session_exists(session_code: str) -> bool:
    return _session_path(session_code).exists()


def delete_session(session_code: str) -> None:
    path = _session_path(session_code)
    if path.exists():
        path.unlink()


def mark_completed(session_code: str) -> None:
    data = load_session(session_code)
    if data:
        data["phase"] = "completed"
        save_session(session_code, data)

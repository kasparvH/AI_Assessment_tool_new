import json
from pathlib import Path

ROOT = Path(__file__).parent
CONFIG_PATH = ROOT / "config.json"

DEFAULTS = {
    "test_mode": False,
    "test_questions_per_dimension": 4,
}


def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return {**DEFAULTS, **json.loads(CONFIG_PATH.read_text(encoding="utf-8"))}
        except Exception:
            pass
    return dict(DEFAULTS)


def save_config(config: dict) -> None:
    CONFIG_PATH.write_text(json.dumps(config, indent=2), encoding="utf-8")


def is_test_mode() -> bool:
    return load_config().get("test_mode", False)


def get_test_questions_per_dim() -> int:
    return load_config().get("test_questions_per_dimension", 4)


def set_test_mode(enabled: bool) -> None:
    cfg = load_config()
    cfg["test_mode"] = enabled
    save_config(cfg)

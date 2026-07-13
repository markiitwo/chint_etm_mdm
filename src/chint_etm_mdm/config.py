from __future__ import annotations

import json
from dataclasses import dataclass, asdict
from pathlib import Path


APP_DIR_NAME = "CHINT ETM MDM"
CONFIG_FILE_NAME = "config.json"


@dataclass
class AppConfig:
    work_dir: str = ""
    db_path: str = ""

    @property
    def work_path(self) -> Path | None:
        return Path(self.work_dir) if self.work_dir else None

    @property
    def database_path(self) -> Path | None:
        return Path(self.db_path) if self.db_path else None


def user_config_path() -> Path:
    base = Path.home() / ".config" / "chint_etm_mdm"
    base.mkdir(parents=True, exist_ok=True)
    return base / CONFIG_FILE_NAME


def load_config() -> AppConfig:
    path = user_config_path()
    if not path.exists():
        return AppConfig()
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return AppConfig()
    return AppConfig(
        work_dir=str(raw.get("work_dir") or ""),
        db_path=str(raw.get("db_path") or ""),
    )


def save_config(config: AppConfig) -> None:
    path = user_config_path()
    path.write_text(
        json.dumps(asdict(config), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def ensure_work_dirs(work_dir: Path) -> None:
    for name in ["database", "input", "output", "downloads/price", "logs", "reports"]:
        (work_dir / name).mkdir(parents=True, exist_ok=True)


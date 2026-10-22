from __future__ import annotations

import json
from dataclasses import dataclass, asdict, replace
from pathlib import Path

from .atomic_files import atomic_write_json


APP_DIR_NAME = "CHINT ETM MDM"
CONFIG_FILE_NAME = "config.json"


@dataclass
class AppConfig:
    work_dir: str = ""
    db_path: str = ""
    class_templates_dir: str = ""
    codes_file: str = ""
    last_new_since: str = ""
    inbox_dir: str = ""  # where incoming files land: the RDP shared folder now, the company share later

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
        class_templates_dir=str(raw.get("class_templates_dir") or ""),
        codes_file=str(raw.get("codes_file") or ""),
        last_new_since=str(raw.get("last_new_since") or ""),
        inbox_dir=str(raw.get("inbox_dir") or ""),
    )


DEFAULT_WORK_DIR_NAME = "CHINT ETM MDM Work"
DB_FILE_NAME = "chint_mdm.sqlite"


def first_run_defaults(config: AppConfig, home: Path | None = None) -> AppConfig:
    """Fill an empty work folder and database path so a first start needs no setup."""
    home = home or Path.home()
    work_dir = config.work_dir or str(home / DEFAULT_WORK_DIR_NAME)
    db_path = config.db_path
    if not db_path:
        candidates = [Path(work_dir) / "database" / DB_FILE_NAME] + [
            home / folder / DB_FILE_NAME for folder in ("Downloads", "Загрузки", "Documents", "Документы", "")
        ]
        db_path = next((str(path) for path in candidates if path.is_file()), "")
    return replace(config, work_dir=work_dir, db_path=db_path)


def incoming_folders(inbox_dir: str = "", home: Path | None = None) -> list[Path]:
    """Where new files from colleagues and partners are looked for: the inbox folder first, then Downloads."""
    home = home or Path.home()
    folders = [Path(inbox_dir)] if inbox_dir else []
    return folders + [home / "Downloads", home / "Загрузки"]


def save_config(config: AppConfig) -> None:
    path = user_config_path()
    atomic_write_json(path, asdict(config))


def ensure_work_dirs(work_dir: Path) -> None:
    for name in ["database", "input", "output", "downloads/price", "logs", "reports", "rules", "etm_class_templates"]:
        (work_dir / name).mkdir(parents=True, exist_ok=True)

from __future__ import annotations

import datetime as dt
import json
import shutil
from dataclasses import dataclass, asdict, replace
from pathlib import Path

from .atomic_files import atomic_write_json


APP_DIR_NAME = "CHINT ETM MDM"
CONFIG_FILE_NAME = "config.json"

# Settings files that could not be read at start; the window tells the user about them.
LOAD_PROBLEMS: list[str] = []


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


def read_settings_file(path: Path, what: str) -> dict:
    """The file's settings, or {} when it is missing or broken.

    A broken file is copied aside before the program can overwrite it with defaults, so the old
    settings can still be looked up.
    """
    if not path.exists():
        return {}
    problem = ""
    try:
        raw = json.loads(path.read_text(encoding="utf-8-sig"))
    except OSError as exc:
        LOAD_PROBLEMS.append(f"{what}: не удалось прочитать {path} ({exc}). Работаю со стандартными значениями.")
        return {}
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raw, problem = {}, f"файл повреждён ({exc})"
    if not isinstance(raw, dict):
        problem = "внутри не настройки, а " + ("пустое значение" if raw is None else type(raw).__name__)
    if not problem:
        return raw
    copy = path.with_name(f"{path.name}.broken-{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}")
    try:
        shutil.copy2(path, copy)
        kept = f" Копия старого файла: {copy}."
    except OSError:
        kept = ""
    LOAD_PROBLEMS.append(f"{what}: {problem}. Работаю со стандартными значениями.{kept}")
    return {}


def text_setting(raw: dict, key: str) -> str:
    value = raw.get(key)
    return value if isinstance(value, str) else ""


def load_config() -> AppConfig:
    raw = read_settings_file(user_config_path(), "Настройки программы")
    return AppConfig(
        work_dir=text_setting(raw, "work_dir"),
        db_path=text_setting(raw, "db_path"),
        class_templates_dir=text_setting(raw, "class_templates_dir"),
        codes_file=text_setting(raw, "codes_file"),
        last_new_since=text_setting(raw, "last_new_since"),
        inbox_dir=text_setting(raw, "inbox_dir"),
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

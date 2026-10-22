from __future__ import annotations

import datetime as dt
import hashlib
import os
import sqlite3
import tempfile
from collections.abc import Iterable
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path

# Tables every CHINT ETM MDM database has; a file without them is not a backup of this program's base.
CHINT_TABLES = ("products", "price_snapshots", "price_snapshot_items")


@dataclass(frozen=True)
class BackupResult:
    path: Path
    sha256: str
    created_at: dt.datetime
    integrity_ok: bool


def connect_database(path: Path, *, readonly: bool = False) -> sqlite3.Connection:
    if readonly:
        conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True, timeout=10)
    else:
        conn = sqlite3.connect(path, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA busy_timeout = 10000")
    return conn


def check_sqlite_integrity(db_path: Path) -> None:
    with closing(connect_database(db_path, readonly=True)) as conn:
        result = str(conn.execute("PRAGMA integrity_check").fetchone()[0])
    if result.lower() != "ok":
        raise ValueError(f"SQLite integrity_check failed for {db_path}: {result}")


class BackupRejected(ValueError):
    """The backup was refused before anything was touched: the base is as it was."""


def missing_tables(db_path: Path, required: Iterable[str] = CHINT_TABLES) -> list[str]:
    """The program's tables a SQLite file lacks; raises sqlite3.DatabaseError for a file that is not SQLite."""
    with closing(connect_database(db_path, readonly=True)) as conn:
        tables = {str(row[0]) for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
    return [name for name in required if name not in tables]


def base_problem(db_path: Path) -> str:
    """Why this file cannot be the program's base, in plain words; "" when it can."""
    try:
        missing = missing_tables(db_path)
    except sqlite3.DatabaseError:
        return "Файл не похож на базу SQLite или повреждён."
    if missing:
        return "Это не база программы: в ней нет таблиц " + ", ".join(missing) + "."
    return ""


def check_backup_file(backup_path: Path, db_path: Path, required_tables: Iterable[str] = ()) -> None:
    """Refuse a backup that would break the base: the base itself, not SQLite, damaged or not this program's."""
    try:
        if backup_path.resolve() == db_path.resolve() or backup_path.samefile(db_path):
            raise BackupRejected("Выбран сам файл базы, а не бэкап. Выберите файл бэкапа (.bak_...).")
    except OSError:
        pass
    try:
        with closing(connect_database(backup_path, readonly=True)) as conn:
            tables = {str(row[0]) for row in conn.execute("SELECT name FROM sqlite_master WHERE type = 'table'")}
            result = str(conn.execute("PRAGMA integrity_check").fetchone()[0])
    except sqlite3.DatabaseError as exc:
        raise BackupRejected(f"Файл не похож на базу SQLite или повреждён: {backup_path}\n({exc})") from exc
    if result.lower() != "ok":
        raise BackupRejected(f"Бэкап повреждён, восстанавливать из него нельзя: {backup_path}\n({result})")
    missing = [name for name in required_tables if name not in tables]
    if missing:
        raise BackupRejected(
            "Это не бэкап базы программы: в файле нет таблиц " + ", ".join(missing) + f".\n{backup_path}"
        )


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def remove_sqlite_sidecars(db_path: Path) -> None:
    for suffix in ("-wal", "-shm"):
        sidecar = Path(f"{db_path}{suffix}")
        try:
            sidecar.unlink(missing_ok=True)
        except OSError as exc:
            raise OSError(f"Could not remove SQLite sidecar file {sidecar}: {exc}") from exc


def create_sqlite_backup(db_path: Path, backup_dir: Path | None = None, label: str = "bak") -> BackupResult:
    if not db_path.exists():
        raise FileNotFoundError(f"SQLite database not found: {db_path}")
    backup_dir = backup_dir or db_path.parent
    backup_dir.mkdir(parents=True, exist_ok=True)
    created_at = dt.datetime.now(dt.timezone.utc)
    stamp = created_at.strftime("%Y%m%dT%H%M%S.%fZ")
    temp_name = ""
    try:
        with tempfile.NamedTemporaryFile(
            delete=False,
            dir=backup_dir,
            prefix=f".{db_path.name}.{label}.",
            suffix=".tmp",
        ) as temp_file:
            temp_name = temp_file.name
        temp_path = Path(temp_name)
        with closing(connect_database(db_path, readonly=True)) as source, closing(
            sqlite3.connect(temp_path)
        ) as dest:
            with source, dest:
                source.backup(dest)
        check_sqlite_integrity(temp_path)
        digest = sha256_file(temp_path)
        final_path = backup_dir / f"{db_path.name}.{label}_{stamp}_{digest[:8]}"
        os.replace(temp_path, final_path)
        temp_name = ""
        return BackupResult(final_path, digest, created_at, True)
    finally:
        if temp_name:
            try:
                Path(temp_name).unlink(missing_ok=True)
            except OSError:
                pass


def copy_sqlite_database(source_path: Path, target_path: Path) -> None:
    if not source_path.exists():
        raise FileNotFoundError(f"SQLite database not found: {source_path}")
    target_path.parent.mkdir(parents=True, exist_ok=True)
    temp_name = ""
    try:
        with tempfile.NamedTemporaryFile(
            delete=False,
            dir=target_path.parent,
            prefix=f".{target_path.name}.copy.",
            suffix=".tmp",
        ) as temp_file:
            temp_name = temp_file.name
        temp_path = Path(temp_name)
        with closing(connect_database(source_path, readonly=True)) as source, closing(
            sqlite3.connect(temp_path)
        ) as dest:
            with source, dest:
                source.backup(dest)
        check_sqlite_integrity(temp_path)
        os.replace(temp_path, target_path)
        temp_name = ""
    finally:
        if temp_name:
            try:
                Path(temp_name).unlink(missing_ok=True)
            except OSError:
                pass


def restore_sqlite_backup(
    db_path: Path, backup_path: Path, required_tables: Iterable[str] = ()
) -> BackupResult:
    if not db_path.exists():
        raise FileNotFoundError(f"Current database not found: {db_path}")
    if not backup_path.exists():
        raise FileNotFoundError(f"Backup not found: {backup_path}")
    check_backup_file(backup_path, db_path, required_tables)
    safety_backup = create_sqlite_backup(db_path, db_path.parent, label="before_restore")
    temp_name = ""
    try:
        with tempfile.NamedTemporaryFile(
            delete=False,
            dir=db_path.parent,
            prefix=f".{db_path.name}.restore.",
            suffix=".tmp",
        ) as temp_file:
            temp_name = temp_file.name
        temp_path = Path(temp_name)
        with closing(connect_database(backup_path, readonly=True)) as source, closing(
            sqlite3.connect(temp_path)
        ) as dest:
            with source, dest:
                source.backup(dest)
        check_sqlite_integrity(temp_path)
        remove_sqlite_sidecars(db_path)
        os.replace(temp_path, db_path)
        temp_name = ""
        check_sqlite_integrity(db_path)
        return safety_backup
    finally:
        if temp_name:
            try:
                Path(temp_name).unlink(missing_ok=True)
            except OSError:
                pass

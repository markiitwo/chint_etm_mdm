from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

from chint_etm_mdm.db_utils import (
    copy_sqlite_database,
    create_sqlite_backup,
    restore_sqlite_backup,
)


class SqliteBackupTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.db_path = self.root / "source.sqlite"
        with closing(sqlite3.connect(self.db_path)) as conn:
            with conn:
                conn.execute("CREATE TABLE sample (value TEXT NOT NULL)")
                conn.execute("INSERT INTO sample (value) VALUES ('original')")

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    @staticmethod
    def read_value(db_path: Path) -> str:
        with closing(sqlite3.connect(db_path)) as conn:
            row = conn.execute("SELECT value FROM sample").fetchone()
        assert row is not None
        return str(row[0])

    def assert_no_temporary_databases(self) -> None:
        self.assertEqual([], list(self.root.rglob(".*.tmp")))

    def test_create_backup_closes_and_renames_temporary_database(self) -> None:
        result = create_sqlite_backup(self.db_path)

        self.assertTrue(result.path.exists())
        self.assertTrue(result.integrity_ok)
        self.assertEqual("original", self.read_value(result.path))
        self.assert_no_temporary_databases()

    def test_copy_database_closes_and_renames_temporary_database(self) -> None:
        target_path = self.root / "copies" / "copy.sqlite"

        copy_sqlite_database(self.db_path, target_path)

        self.assertEqual("original", self.read_value(target_path))
        self.assert_no_temporary_databases()

    def test_restore_database_preserves_current_database_as_safety_backup(self) -> None:
        backup = create_sqlite_backup(self.db_path)
        with closing(sqlite3.connect(self.db_path)) as conn:
            with conn:
                conn.execute("UPDATE sample SET value = 'changed'")

        safety_backup = restore_sqlite_backup(self.db_path, backup.path)

        self.assertEqual("original", self.read_value(self.db_path))
        self.assertEqual("changed", self.read_value(safety_backup.path))
        self.assert_no_temporary_databases()


if __name__ == "__main__":
    unittest.main()

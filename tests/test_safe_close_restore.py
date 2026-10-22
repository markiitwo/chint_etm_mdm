"""Closing the window during an operation, and rolling the base back only from a real backup."""

from __future__ import annotations

import os
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path
from unittest import mock

from chint_etm_mdm.db_utils import CHINT_TABLES, create_sqlite_backup, restore_sqlite_backup


def make_base(path: Path, name: str) -> None:
    with closing(sqlite3.connect(path)) as conn, conn:
        for table in CHINT_TABLES:
            conn.execute(f"CREATE TABLE {table} (name TEXT)")
        conn.execute("INSERT INTO products (name) VALUES (?)", (name,))


def product_name(path: Path) -> str:
    with closing(sqlite3.connect(path)) as conn:
        return str(conn.execute("SELECT name FROM products").fetchone()[0])


class RestoreChecksTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.db = self.root / "chint_mdm.sqlite"
        make_base(self.db, "current")

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def assert_untouched(self) -> None:
        self.assertEqual("current", product_name(self.db))
        # No safety copy either: nothing was started.
        self.assertEqual([], list(self.root.glob("*before_restore*")))

    def test_restores_a_backup_of_this_program(self) -> None:
        backup = self.root / "old.bak"
        make_base(backup, "old")
        safety = restore_sqlite_backup(self.db, backup, CHINT_TABLES)
        self.assertEqual("old", product_name(self.db))
        self.assertEqual("current", product_name(safety.path))

    def test_refuses_the_base_itself(self) -> None:
        with self.assertRaisesRegex(ValueError, "сам файл базы"):
            restore_sqlite_backup(self.db, self.db, CHINT_TABLES)
        self.assert_untouched()

    def test_refuses_a_file_that_is_not_sqlite(self) -> None:
        not_a_base = self.root / "price.xlsx"
        not_a_base.write_bytes(b"PK\x03\x04 not a database" * 100)
        with self.assertRaisesRegex(ValueError, "не похож на базу SQLite"):
            restore_sqlite_backup(self.db, not_a_base, CHINT_TABLES)
        self.assert_untouched()

    def test_refuses_another_programs_base(self) -> None:
        other = self.root / "other.sqlite"
        with closing(sqlite3.connect(other)) as conn, conn:
            conn.execute("CREATE TABLE products (name TEXT)")
        with self.assertRaisesRegex(ValueError, "price_snapshots"):
            restore_sqlite_backup(self.db, other, CHINT_TABLES)
        self.assert_untouched()

    def test_restore_worker_reports_plainly_and_keeps_the_base(self) -> None:
        from chint_etm_mdm.app import RestoreWorker

        failures: list[str] = []
        worker = RestoreWorker(self.db, self.db)
        worker.failed.connect(failures.append)
        worker.done.connect(lambda _: self.fail("restore must not run"))
        worker.run()
        self.assertTrue(failures[0].startswith("База не изменена."))
        self.assertNotIn("Traceback", failures[0])
        self.assert_untouched()

    def test_restore_worker_returns_the_safety_copy(self) -> None:
        from chint_etm_mdm.app import RestoreWorker

        backup = create_sqlite_backup(self.db).path
        with closing(sqlite3.connect(self.db)) as conn, conn:
            conn.execute("UPDATE products SET name = 'changed'")
        results: list[Path] = []
        worker = RestoreWorker(self.db, backup)
        worker.done.connect(results.append)
        worker.failed.connect(self.fail)
        worker.run()
        self.assertEqual("current", product_name(self.db))
        self.assertEqual("changed", product_name(results[0]))


class FakeRunningWorker:
    def __init__(self) -> None:
        self.running = True

    def isRunning(self) -> bool:
        return self.running


class CloseDuringOperationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
        from PySide6.QtWidgets import QApplication

        cls.app = QApplication.instance() or QApplication([])

    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        home = Path(self.temp_dir.name)
        # The window reads and writes its settings in the home folder: keep the real ones out of it.
        self.env = mock.patch.dict(os.environ, {"HOME": str(home), "USERPROFILE": str(home), "APPDATA": str(home)})
        self.env.start()
        self.home = mock.patch("pathlib.Path.home", return_value=home)
        self.home.start()
        from chint_etm_mdm.app import MainWindow

        self.window = MainWindow()

    def tearDown(self) -> None:
        self.window.worker = None
        self.window.close()
        self.home.stop()
        self.env.stop()
        self.temp_dir.cleanup()

    def close_with_answer(self, answer):
        from PySide6.QtGui import QCloseEvent
        from PySide6.QtWidgets import QMessageBox

        event = QCloseEvent()
        with mock.patch.object(QMessageBox, "question", return_value=answer) as asked:
            self.window.closeEvent(event)
        return event, asked

    def test_closes_at_once_when_idle(self) -> None:
        from PySide6.QtWidgets import QMessageBox

        event, asked = self.close_with_answer(QMessageBox.StandardButton.Yes)
        self.assertTrue(event.isAccepted())
        asked.assert_not_called()

    def test_stays_open_when_user_says_no(self) -> None:
        from PySide6.QtWidgets import QMessageBox

        self.window.worker = FakeRunningWorker()
        event, asked = self.close_with_answer(QMessageBox.StandardButton.No)
        self.assertFalse(event.isAccepted())
        asked.assert_called_once()
        self.assertFalse(self.window.close_when_done)

    def test_closes_itself_after_the_operation(self) -> None:
        from PySide6.QtWidgets import QMessageBox

        worker = FakeRunningWorker()
        self.window.worker = worker
        event, _ = self.close_with_answer(QMessageBox.StandardButton.Yes)
        self.assertFalse(event.isAccepted())
        self.assertTrue(self.window.close_when_done)

        # A second close while waiting does not ask again.
        event, asked = self.close_with_answer(QMessageBox.StandardButton.No)
        self.assertFalse(event.isAccepted())
        asked.assert_not_called()

        with mock.patch("chint_etm_mdm.app.QTimer.singleShot") as later:
            worker.running = False
            self.window.on_worker_finished(worker)
        later.assert_called_once_with(0, self.window.close)
        self.assertIsNone(self.window.worker)

    def test_finishing_without_a_close_request_keeps_the_window(self) -> None:
        worker = FakeRunningWorker()
        self.window.worker = worker
        with mock.patch("chint_etm_mdm.app.QTimer.singleShot") as later:
            self.window.on_worker_finished(worker)
        later.assert_not_called()


if __name__ == "__main__":
    unittest.main()

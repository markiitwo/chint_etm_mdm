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
        from PySide6.QtWidgets import QMessageBox

        from chint_etm_mdm.app import MainWindow

        # A dialog nobody answers would hang the test run: by default every dialog returns at once.
        self.dialogs = [
            mock.patch.object(QMessageBox, name, return_value=QMessageBox.StandardButton.No)
            for name in ("warning", "information", "critical", "question")
        ]
        for dialog in self.dialogs:
            dialog.start()
        self.window = MainWindow()

    def tearDown(self) -> None:
        self.window.worker = None
        self.window.close()
        for dialog in self.dialogs:
            dialog.stop()
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


class ChosenPathsTests(CloseDuringOperationTests):
    """GUI-01/GUI-03: plain messages for empty paths and folders; old ETIM reports can be picked."""

    def warned(self, text: str):
        from PySide6.QtWidgets import QMessageBox

        with mock.patch.object(QMessageBox, "warning") as warning:
            result = self.window.chosen_file(text, "База", "SQLite базу")
        return result, warning

    def test_empty_field_is_not_the_current_folder(self) -> None:
        result, warning = self.warned("  ")
        self.assertIsNone(result)
        self.assertIn("не выбрано", warning.call_args[0][1])

    def test_folder_instead_of_file(self) -> None:
        result, warning = self.warned(self.temp_dir.name)
        self.assertIsNone(result)
        self.assertIn("выбрана папка", warning.call_args[0][1])

    def test_missing_and_existing_file(self) -> None:
        missing, warning = self.warned(str(Path(self.temp_dir.name) / "нет.sqlite"))
        self.assertIsNone(missing)
        self.assertIn("не найден", warning.call_args[0][1])
        base = Path(self.temp_dir.name) / "base.sqlite"
        base.write_bytes(b"")
        found, warning = self.warned(str(base))
        self.assertEqual(base, found)
        warning.assert_not_called()

    def test_apply_etim_button_works_after_restart(self) -> None:
        self.assertIsNone(self.window.last_etim_report_path)
        self.window.update_write_action_buttons()
        self.assertTrue(self.window.apply_etim_report_button.isEnabled())
        self.window.worker = FakeRunningWorker()
        self.window.update_write_action_buttons()
        self.assertFalse(self.window.apply_etim_report_button.isEnabled())

    def test_apply_etim_always_offers_a_choice(self) -> None:
        base = Path(self.temp_dir.name) / "base.sqlite"
        make_base(base, "x")
        self.window.db_path_edit.setText(str(base))
        last = Path(self.temp_dir.name) / "last.xlsx"
        last.write_bytes(b"")
        self.window.last_etim_report_path = last
        with mock.patch.object(self.window, "choose_etim_report_file", return_value=None) as choose:
            self.window.apply_etim_decisions()
        choose.assert_called_once_with(last)


class RulesDuringOperationTests(CloseDuringOperationTests):
    def test_rules_do_not_change_under_a_running_operation(self) -> None:
        from PySide6.QtWidgets import QMessageBox

        self.window.worker = FakeRunningWorker()
        row = {"class81_code": "C1", "template_field": "Ток", "source_attribute": "A"}
        with mock.patch.object(QMessageBox, "warning") as warning, mock.patch(
            "chint_etm_mdm.app.add_approved_class_rule"
        ) as approve, mock.patch("chint_etm_mdm.app.add_rejected_class_rule") as reject:
            self.assertFalse(self.window.save_rule_data(row))
            self.assertFalse(self.window.reject_rule_data(row))
            self.window.save_selected_rules()
            self.window.reject_selected_rules()
        approve.assert_not_called()
        reject.assert_not_called()
        self.assertEqual(4, warning.call_count)


class EtmTabTests(CloseDuringOperationTests):
    def test_typed_articles_and_again(self) -> None:
        base = Path(self.temp_dir.name) / "base.sqlite"
        make_base(base, "x")
        self.window.db_path_edit.setText(str(base))
        self.window.class_templates_edit.setText(str(Path(self.temp_dir.name) / "templates"))
        self.window.etm_articles_edit.setPlainText("586772, 625288")
        self.assertTrue(self.window.etm_list_mode.isChecked())
        with mock.patch.object(self.window, "start_worker") as start:
            self.window.run_new_products()
        worker = start.call_args[0][0]
        self.assertEqual(["586772", "625288"], worker.articles)
        self.assertEqual("ЭТМ", self.window.tabs.tabText(1))

        self.window.config.last_etm_articles = "111 222"
        with mock.patch.object(self.window, "start_worker") as start:
            self.window.run_etm_again()
        self.assertEqual(["111", "222"], start.call_args[0][0].articles)

        self.window.etm_new_mode.setChecked(True)
        with mock.patch.object(self.window, "start_worker") as start:
            self.window.etm_run_button.click()  # through the button, as the user does
        self.assertIsNone(start.call_args[0][0].articles)


class EtmRequestsAfterRestartTests(CloseDuringOperationTests):
    def test_saved_needs_of_the_newest_run(self) -> None:
        import json

        work = Path(self.temp_dir.name) / "work"
        self.window.work_dir_edit.setText(str(work))
        self.assertIsNone(self.window.saved_etm_needs())
        tech = work / "output" / "ЭТМ карточки 2026-10-22 10-00" / "Служебное"
        tech.mkdir(parents=True)
        need = {"article": "1", "name": "n", "series": "s", "key": "Код ТН ВЭД", "label": "Код ТН ВЭД", "kind": "", "allowed": []}
        (tech / "запросы_продактам.json").write_text(json.dumps({"version": 1, "needs": [need]}), encoding="utf-8")
        folder, needs = self.window.saved_etm_needs()
        self.assertEqual("ЭТМ карточки 2026-10-22 10-00", folder.name)
        with mock.patch.object(self.window, "ask_managers") as ask:
            self.window.write_etm_pm_requests()
        self.assertEqual(["1"], [item.article for item in ask.call_args[0][0]])


class WriteButtonsTests(CloseDuringOperationTests):
    def test_buttons_of_every_tab_wait_for_the_operation(self) -> None:
        texts = {button.text() for button in self.window.write_action_buttons}
        for text in ("Заполнить архив РАЭК", "Сделать материалы", "Откатить базу из бэкапа", "Импортировать ETIM"):
            self.assertIn(text, texts)
        self.window.worker = FakeRunningWorker()
        self.window.update_write_action_buttons()
        self.assertFalse(any(button.isEnabled() for button in self.window.write_action_buttons))


class BaseProblemTests(unittest.TestCase):
    def test_plain_reasons(self) -> None:
        from chint_etm_mdm.db_utils import base_problem

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            good = root / "good.sqlite"
            make_base(good, "x")
            self.assertEqual("", base_problem(good))
            empty = root / "empty.sqlite"
            empty.write_bytes(b"")
            self.assertIn("не база программы", base_problem(empty))
            text = root / "price.sqlite"
            text.write_bytes(b"PK\x03\x04 not sqlite" * 50)
            self.assertIn("не похож на базу SQLite", base_problem(text))

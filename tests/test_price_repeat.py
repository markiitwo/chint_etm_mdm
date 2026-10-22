"""PRI-04: the same price twice, "_1" copies, duplicate rows and 100 vs 100.00."""

from __future__ import annotations

import shutil
import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

from chint_etm_mdm.price_importer import (
    already_imported,
    import_price_workbook,
    result_lines,
    snapshot_date_from_filename,
)
from test_import_data_integrity import make_database, write_price

HEADER = ["Артикул", "Наименование", "Тариф с НДС, руб", "Тариф без НДС, руб"]


def count(db: Path, query: str) -> int:
    with closing(sqlite3.connect(db)) as conn:
        return int(conn.execute(query).fetchone()[0])


class SnapshotDateTests(unittest.TestCase):
    def test_copy_suffix_is_not_part_of_the_date(self) -> None:
        for name, date in [
            ("Price-list-CHINT_01.10.2026.xlsx", "01.10.2026"),
            ("Price-list-CHINT_01.10.2026_1.xlsx", "01.10.2026"),
            ("Price-list-CHINT_01.10.2026_12.xlsx", "01.10.2026"),
            ("Price-list-CHINT_2026-10-01 (2).xlsx", "2026-10-01"),
            ("Price-list-CHINT_2026_10_01.xlsx", "2026_10_01"),
            ("my_price_1.xlsx", "my_price_1"),  # no date: the name stays as it was
        ]:
            with self.subTest(name=name):
                self.assertEqual(date, snapshot_date_from_filename(Path(name)))


class RepeatImportTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.db = self.root / "db.sqlite"
        make_database(self.db)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def price(self, name: str, rows) -> Path:
        return write_price(self.root / name, HEADER, rows)

    def test_same_file_again_changes_nothing(self) -> None:
        first = self.price("Price-list-CHINT_01.10.2026.xlsx", [["A-1", "Автомат", "100", "83,33"]])
        import_price_workbook(first, self.db, make_backup=False)
        copy = self.root / "Price-list-CHINT_01.10.2026_1.xlsx"
        shutil.copy2(first, copy)
        self.assertIsNotNone(already_imported(copy, self.db))

        again = import_price_workbook(copy, self.db, make_backup=True)
        self.assertEqual(1, again.repeat_of_snapshot)
        self.assertEqual("01.10.2026", again.repeat_of_date)
        self.assertIsNone(again.backup_path)
        self.assertEqual(1, count(self.db, "SELECT COUNT(*) FROM price_snapshots"))
        self.assertEqual([], list(self.root.glob("*.bak_*")))
        self.assertIn("уже был импортирован", "\n".join(result_lines(again)))

        # On purpose (after a rollback, say) it goes in again.
        forced = import_price_workbook(copy, self.db, make_backup=False, allow_repeat=True)
        self.assertIsNone(forced.repeat_of_snapshot)
        self.assertEqual(2, count(self.db, "SELECT COUNT(*) FROM price_snapshots"))
        self.assertEqual("01.10.2026", forced.snapshot_date)

    def test_base_from_before_this_version_imports_normally(self) -> None:
        self.assertIsNone(already_imported(self.price("p.xlsx", [["A-1", "x", "1", "1"]]), self.db))

    def test_same_price_written_differently_is_not_a_change(self) -> None:
        import_price_workbook(self.price("Price-list-CHINT_01.10.2026.xlsx", [["A-1", "x", "100", "83.3"], ["A-2", "y", "50", "40"]]), self.db, make_backup=False)
        result = import_price_workbook(
            self.price("Price-list-CHINT_02.10.2026.xlsx", [["A-1", "x", "100,00", "83,30"], ["A-2", "y", "55", "45"]]),
            self.db,
            make_backup=False,
        )
        self.assertEqual(1, result.changed_price_rows)

    def test_duplicate_rows(self) -> None:
        import_price_workbook(self.price("Price-list-CHINT_01.10.2026.xlsx", [["A-1", "x", "100", "80"]]), self.db, make_backup=False)
        result = import_price_workbook(
            self.price(
                "Price-list-CHINT_02.10.2026.xlsx",
                [["A-1", "x", "120", "100"], ["A-1", "x", "120", "100"], ["B-1", "b", "1", "1"], ["B-1", "b", "2", "2"]],
            ),
            self.db,
            make_backup=False,
        )
        self.assertEqual(3, result.price_rows)  # the identical A-1 row is written once
        self.assertEqual(("B-1",), result.conflicting_articles)
        self.assertEqual(1, result.changed_price_rows)  # A-1 once, not twice
        self.assertTrue(any("B-1" in line for line in result_lines(result)))


if __name__ == "__main__":
    unittest.main()

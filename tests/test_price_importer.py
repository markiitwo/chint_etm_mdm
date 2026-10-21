from __future__ import annotations

import sqlite3
import unittest
from contextlib import closing
from pathlib import Path
from unittest.mock import patch

from chint_etm_mdm.price_importer import PRICE_SHEETS, latest_snapshot_id, read_price_workbook


class PriceWorkbookParsingTests(unittest.TestCase):
    @staticmethod
    def valid_auxiliary_rows() -> list[list[str]]:
        return [["Артикул", "Наименование"], ["A-1", "Товар"]]

    def read_with_rows(self, rows_by_sheet: dict[str, list[list[str]]]) -> dict[str, list[list[str]]]:
        with patch(
            "chint_etm_mdm.price_importer.read_sheet_rows",
            side_effect=lambda _path, sheet_name: rows_by_sheet[sheet_name],
        ):
            return read_price_workbook(Path("price.xlsx"))

    def test_accepts_prepared_workbook_with_header_in_first_row(self) -> None:
        rows_by_sheet = {sheet: self.valid_auxiliary_rows() for sheet in PRICE_SHEETS}

        workbook = self.read_with_rows(rows_by_sheet)

        self.assertEqual(["Артикул", "Наименование"], workbook["Price-list"][0])
        self.assertEqual(["A-1", "Товар"], workbook["Price-list"][1])

    def test_finds_price_header_after_logo_and_metadata_rows(self) -> None:
        rows_by_sheet = {sheet: self.valid_auxiliary_rows() for sheet in PRICE_SHEETS}
        rows_by_sheet["Price-list"] = [
            [],
            ["Дата актуальности: 04.08.2026"],
            ["Дата последнего изменения тарифа: 12.01.2026"],
            ["Артикул ", " Наименование", "Тариф с НДС, руб"],
            ["A-2", "Новый товар", "100"],
        ]

        workbook = self.read_with_rows(rows_by_sheet)

        self.assertEqual(
            ["Артикул", "Наименование", "Тариф с НДС, руб"],
            workbook["Price-list"][0],
        )
        self.assertEqual(["A-2", "Новый товар", "100"], workbook["Price-list"][1])

    def test_rejects_workbook_without_required_headers(self) -> None:
        rows_by_sheet = {sheet: self.valid_auxiliary_rows() for sheet in PRICE_SHEETS}
        rows_by_sheet["Price-list"] = [["Дата актуальности: 04.08.2026"], ["A-1", "Товар"]]

        with self.assertRaisesRegex(ValueError, "обязательные заголовки"):
            self.read_with_rows(rows_by_sheet)

    def test_rejects_price_sheet_without_articles(self) -> None:
        rows_by_sheet = {sheet: self.valid_auxiliary_rows() for sheet in PRICE_SHEETS}
        rows_by_sheet["Price-list"] = [["Артикул", "Наименование"]]

        with self.assertRaisesRegex(ValueError, "не содержит строк с артикулами"):
            self.read_with_rows(rows_by_sheet)


class SnapshotSelectionTests(unittest.TestCase):
    def test_uses_latest_snapshot_with_price_items(self) -> None:
        with closing(sqlite3.connect(":memory:")) as conn:
            conn.executescript(
                """
                CREATE TABLE price_snapshots (id INTEGER PRIMARY KEY);
                CREATE TABLE price_snapshot_items (snapshot_id INTEGER NOT NULL, article TEXT NOT NULL, price_with_vat TEXT);
                INSERT INTO price_snapshots (id) VALUES (1), (2), (3);
                INSERT INTO price_snapshot_items (snapshot_id, article, price_with_vat) VALUES (1, 'A-1', '10'), (2, 'A-2', '20');
                """
            )

            self.assertEqual(2, latest_snapshot_id(conn))

    def test_returns_none_when_all_snapshots_are_empty(self) -> None:
        with closing(sqlite3.connect(":memory:")) as conn:
            conn.executescript(
                """
                CREATE TABLE price_snapshots (id INTEGER PRIMARY KEY);
                CREATE TABLE price_snapshot_items (snapshot_id INTEGER NOT NULL, article TEXT NOT NULL, price_with_vat TEXT);
                INSERT INTO price_snapshots (id) VALUES (1), (2);
                """
            )

            self.assertIsNone(latest_snapshot_id(conn))


if __name__ == "__main__":
    unittest.main()

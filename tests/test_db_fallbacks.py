"""DB-03: empty text does not hide a filled source; the manufacturer is not a unit."""

from __future__ import annotations

import sqlite3
import tempfile
import unittest
from contextlib import closing
from pathlib import Path

from chint_etm_mdm.db import fetch_products
from test_import_data_integrity import make_database


class FallbackTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db = Path(self.temp_dir.name) / "db.sqlite"
        make_database(self.db)

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def run_sql(self, *statements: tuple[str, tuple]) -> None:
        with closing(sqlite3.connect(self.db)) as conn, conn:
            for sql, params in statements:
                conn.execute(sql, params)

    def test_blank_values_fall_through(self) -> None:
        self.run_sql(
            ("INSERT INTO products (article, name, unit) VALUES (?, ?, ?)", ("A-1", "Имя из прайса", "  ")),
            ("INSERT INTO ipro_goods (article, name, full_name, class81_code, manufacturer) VALUES (?, ?, ?, ?, ?)", ("A-1", " ", "", "", "CHINT")),
            ("INSERT INTO product_etm_class_suggestions VALUES (?, ?, ?)", ("A-1", "", 0.9)),
            ("INSERT INTO product_etm_class_suggestions VALUES (?, ?, ?)", ("A-1", "50301005", 0.5)),
            ("INSERT INTO product_okpd2_codes (article, okpd2_9, okpd2_6, import_id) VALUES (?, ?, ?, ?)", ("A-1", " ", "27.12.2", 1)),
        )
        product = fetch_products(self.db, ["A-1"])["A-1"]
        self.assertEqual("Имя из прайса", product.name)
        self.assertEqual("Имя из прайса", product.full_name)
        self.assertEqual("50301005", product.class81_code)
        self.assertEqual("27.12.2", product.okpd2_code)
        self.assertEqual("шт", product.unit)  # not "CHINT"

    def test_filled_values_keep_their_priority(self) -> None:
        self.run_sql(
            ("INSERT INTO products (article, name, unit) VALUES (?, ?, ?)", ("A-1", "Имя из прайса", "компл")),
            ("INSERT INTO ipro_goods (article, name, full_name, class81_code) VALUES (?, ?, ?, ?)", ("A-1", "Имя iPRO", "Полное iPRO", "C-1")),
            ("INSERT INTO product_etm_class_suggestions VALUES (?, ?, ?)", ("A-1", "C-2", 0.9)),
        )
        product = fetch_products(self.db, ["A-1"])["A-1"]
        self.assertEqual(("Имя iPRO", "Полное iPRO", "C-1", "компл"), (product.name, product.full_name, product.class81_code, product.unit))


if __name__ == "__main__":
    unittest.main()

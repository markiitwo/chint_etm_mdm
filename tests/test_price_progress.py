from contextlib import closing, ExitStack
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import patch

from chint_etm_mdm.app import PriceImportWorker
from chint_etm_mdm import price_importer


class PriceProgressTests(unittest.TestCase):
    def test_download_is_reported_even_when_import_fails(self):
        path = Path("downloaded.xlsx")
        worker = PriceImportWorker(Path("db.sqlite"), None, "https://example.test/price.xlsx", Path("downloads"))
        downloaded, progress, errors, results = [], [], [], []
        worker.downloaded.connect(downloaded.append)
        worker.progress.connect(progress.append)
        worker.failed.connect(errors.append)
        worker.done.connect(results.append)

        def fail_import(*args, **kwargs):
            kwargs["progress"]("Создаю резервную копию базы...")
            raise RuntimeError("backup failed")

        with (
            patch("chint_etm_mdm.app.download_price_file", return_value=path),
            patch("chint_etm_mdm.app.import_price_workbook", side_effect=fail_import),
        ):
            worker.run()
        self.assertEqual([path], downloaded)
        self.assertTrue(any("Скачивание завершено" in item for item in progress))
        self.assertIn("резервную копию", progress[-1])
        self.assertEqual([], results)
        self.assertEqual(1, len(errors))

    def test_statistics_failure_rolls_back_snapshot_and_product(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            database, price = root / "db.sqlite", root / "price.xlsx"
            price.touch()
            with closing(sqlite3.connect(database)) as conn, conn:
                conn.executescript("""
                    CREATE TABLE price_snapshots (
                        id INTEGER PRIMARY KEY, snapshot_date TEXT, source_url TEXT,
                        local_file TEXT, workbook_name TEXT);
                    CREATE TABLE products(article TEXT);
                """)
            messages = []

            def insert_product(cursor, snapshot_id, rows):
                cursor.execute("INSERT INTO products VALUES ('new')")
                return 1

            with ExitStack() as stack:
                stack.enter_context(patch.object(price_importer, "read_price_workbook", return_value={s: [] for s in price_importer.PRICE_SHEETS}))
                stack.enter_context(patch.object(price_importer, "latest_snapshot_id", return_value=None))
                stack.enter_context(patch.object(price_importer, "insert_price_rows", side_effect=insert_product))
                for name in ("insert_new_products", "insert_outgoing", "insert_history"):
                    stack.enter_context(patch.object(price_importer, name, return_value=0))
                stack.enter_context(patch.object(price_importer, "count_new_articles", side_effect=RuntimeError("statistics failed")))
                with self.assertRaisesRegex(RuntimeError, "statistics failed"):
                    price_importer.import_price_workbook(price, database, make_backup=False, progress=messages.append)
            self.assertTrue(any("Подсчитываю" in item for item in messages))
            with closing(sqlite3.connect(database)) as conn:
                self.assertEqual(0, conn.execute("SELECT COUNT(*) FROM products").fetchone()[0])
                self.assertEqual(0, conn.execute("SELECT COUNT(*) FROM price_snapshots").fetchone()[0])


if __name__ == "__main__":
    unittest.main()

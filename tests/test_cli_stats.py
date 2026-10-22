"""CLI-01: --stats alone prints the statistics and ends without an argument error."""

from __future__ import annotations

import io
import sqlite3
import sys
import tempfile
import unittest
from contextlib import closing, redirect_stdout
from pathlib import Path
from unittest import mock

from chint_etm_mdm import cli
from test_import_data_integrity import make_database


class StatsOnlyTests(unittest.TestCase):
    def test_stats_alone_succeeds(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db = Path(tmp) / "db.sqlite"
            make_database(db)
            with closing(sqlite3.connect(db)) as conn, conn:
                conn.execute("INSERT INTO products (article) VALUES ('A-1')")
            out = io.StringIO()
            with mock.patch.object(sys, "argv", ["cli", "--db", str(db), "--stats"]), redirect_stdout(out):
                cli.main()
            self.assertIn("products=1", out.getvalue())

    def test_fill_without_template_is_still_an_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            db = Path(tmp) / "db.sqlite"
            make_database(db)
            with mock.patch.object(sys, "argv", ["cli", "--db", str(db)]), mock.patch("sys.stderr", io.StringIO()):
                with self.assertRaises(SystemExit) as stop:
                    cli.main()
            self.assertNotEqual(0, stop.exception.code)


if __name__ == "__main__":
    unittest.main()

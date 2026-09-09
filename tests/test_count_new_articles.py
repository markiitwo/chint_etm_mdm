from __future__ import annotations

import sqlite3
import unittest

from chint_etm_mdm.price_importer import count_new_articles


class CountNewArticlesTests(unittest.TestCase):
    def setUp(self) -> None:
        self.connection = sqlite3.connect(":memory:")
        self.connection.execute(
            "CREATE TABLE price_snapshot_items (snapshot_id INTEGER NOT NULL, article TEXT)"
        )

    def tearDown(self) -> None:
        self.connection.close()

    def add(self, snapshot_id: int, *articles: str | None) -> None:
        self.connection.executemany(
            "INSERT INTO price_snapshot_items (snapshot_id, article) VALUES (?, ?)",
            ((snapshot_id, article) for article in articles),
        )

    def test_first_snapshot_counts_distinct_non_null_articles(self) -> None:
        self.add(1, "A", "A", "B", None)

        self.assertEqual(2, count_new_articles(self.connection, None, 1))

    def test_comparison_counts_only_distinct_articles_missing_from_previous(self) -> None:
        self.add(1, "A", "B", "B", "OLD")
        self.add(2, "A", "B", "C", "C", "D")

        self.assertEqual(2, count_new_articles(self.connection, 1, 2))

    def test_null_is_not_counted_as_new(self) -> None:
        self.add(1, "A")
        self.add(2, None, None)

        self.assertEqual(0, count_new_articles(self.connection, 1, 2))

    def test_empty_current_snapshot_has_no_new_articles(self) -> None:
        self.add(1, "A", "B")

        self.assertEqual(0, count_new_articles(self.connection, 1, 2))

    def test_empty_previous_snapshot_counts_current_distinct_articles(self) -> None:
        self.add(2, "A", "A", "B", None)

        self.assertEqual(2, count_new_articles(self.connection, 1, 2))


if __name__ == "__main__":
    unittest.main()

"""MAP-01/02: a rejected source stops working, approval brings it back, edits by hand are seen."""

from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.mapping_rules import (
    add_approved_class_rule,
    add_rejected_class_rule,
    source_attributes_for,
)


class RejectApprovedRuleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.rules = Path(self.temp_dir.name) / "attribute_mappings.json"
        self.rules.write_text(json.dumps({"version": 1, "class_rules": [], "rejected_class_rules": []}), encoding="utf-8")

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def sources(self) -> tuple[str, ...]:
        return source_attributes_for("C1", "Ток", self.rules)

    def test_approve_reject_approve(self) -> None:
        add_approved_class_rule(self.rules, "C1", "Ток", "Ном. ток")
        add_approved_class_rule(self.rules, "C1", "Ток", "Ток, А")
        self.assertEqual(("Ном. ток", "Ток, А"), self.sources())

        self.assertTrue(add_rejected_class_rule(self.rules, "C1", "Ток", "Ном. ток"))
        self.assertEqual(("Ток, А",), self.sources())  # the other approved source stays
        self.assertFalse(add_rejected_class_rule(self.rules, "C1", "Ток", "Ном. ток"))  # nothing new

        self.assertTrue(add_approved_class_rule(self.rules, "C1", "Ток", "Ном. ток"))
        self.assertEqual(("Ток, А", "Ном. ток"), self.sources())
        data = json.loads(self.rules.read_text(encoding="utf-8"))
        self.assertEqual([], data["rejected_class_rules"])

    def test_rejection_wins_over_an_approval_left_in_the_file(self) -> None:
        self.rules.write_text(
            json.dumps(
                {
                    "class_rules": [{"class81_code": "C1", "template_field": "Ток", "source_attributes": ["A"], "confidence": "approved_class_rule"}],
                    "rejected_class_rules": [{"class81_code": "C1", "template_field": "Ток", "source_attribute": "A"}],
                }
            ),
            encoding="utf-8",
        )
        self.assertEqual((), self.sources())

    def test_file_changed_by_hand_is_picked_up(self) -> None:
        add_approved_class_rule(self.rules, "C1", "Ток", "A")
        self.assertEqual(("A",), self.sources())
        data = json.loads(self.rules.read_text(encoding="utf-8"))
        data["class_rules"][0]["source_attributes"] = ["B-длинное имя"]
        self.rules.write_text(json.dumps(data), encoding="utf-8")
        stat = self.rules.stat()
        os.utime(self.rules, ns=(stat.st_atime_ns, stat.st_mtime_ns + 1_000_000_000))
        self.assertEqual(("B-длинное имя",), self.sources())

    def test_broken_entry_gives_a_plain_error(self) -> None:
        self.rules.write_text(json.dumps({"class_rules": ["oops"]}), encoding="utf-8")
        with self.assertRaisesRegex(ValueError, "правило №1"):
            self.sources()


if __name__ == "__main__":
    unittest.main()

import json
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import etm_golden_cases  # noqa: E402


class EtmGoldenOutputTests(unittest.TestCase):
    """ETM upload_goods output must stay byte-for-byte the same as the recorded snapshot."""

    maxDiff = None

    @classmethod
    def setUpClass(cls) -> None:
        cls.expected = json.loads(etm_golden_cases.FIXTURE_PATH.read_text(encoding="utf-8"))
        cls.actual = etm_golden_cases.normalize(etm_golden_cases.build_snapshot())

    def test_product_values(self) -> None:
        self.assertEqual(self.expected["product_value"], self.actual["product_value"])

    def test_header_rules(self) -> None:
        self.assertEqual(self.expected["headers"], self.actual["headers"])

    def test_xlsx_output(self) -> None:
        self.assertEqual(self.expected["xlsx"], self.actual["xlsx"])

    def test_csv_output(self) -> None:
        self.assertEqual(self.expected["csv"], self.actual["csv"])


if __name__ == "__main__":
    unittest.main()

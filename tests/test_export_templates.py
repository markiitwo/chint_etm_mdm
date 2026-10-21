import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.db import ProductRecord
from chint_etm_mdm.export_templates import available_partners, load_partner_template, parse_partner_template
from chint_etm_mdm.filler import fill_template


def minimal_template(**overrides):
    data = {
        "code": "test",
        "writer": {"kind": "fill_template", "formats": ["csv"]},
        "article_headers": ["SKU"],
        "fields": [
            {"header": "Brand", "const": "CHINT", "source": "static"},
            {"header": "Weight", "field": "weight_kg", "transform": "decimal", "digits": 2, "source": "db"},
        ],
    }
    data.update(overrides)
    return data


class PartnerTemplateTests(unittest.TestCase):
    def test_etm_template_is_bundled(self) -> None:
        self.assertIn("etm", available_partners())
        self.assertEqual("ЭТМ", load_partner_template("etm").name)

    def test_unknown_partner_is_reported(self) -> None:
        with self.assertRaises(ValueError):
            load_partner_template("no-such-partner")

    def test_new_partner_needs_only_json(self) -> None:
        template = parse_partner_template(minimal_template())
        product = ProductRecord(article="A-1", weight_kg=1.234)
        self.assertEqual(("CHINT", "filled", "static"), template.product_value(product, "Brand"))
        self.assertEqual(("1.23", "filled", "db"), template.product_value(product, "Weight"))
        self.assertEqual(("", "blank", ""), template.product_value(product, "Other"))
        self.assertEqual(0, template.find_article_index(["SKU", "Brand"]))

    def test_typo_in_product_field_is_rejected(self) -> None:
        bad = minimal_template(fields=[{"header": "Weight", "field": "weigth_kg"}])
        with self.assertRaises(ValueError):
            parse_partner_template(bad)

    def test_unknown_note_is_rejected(self) -> None:
        bad = minimal_template(fields=[{"header": "Weight", "field": "weight_kg", "note": "nope"}])
        with self.assertRaises(ValueError):
            parse_partner_template(bad)

    def test_unsupported_file_type_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            template = Path(tmp) / "upload_goods.txt"
            template.write_text("Артикул\n", encoding="utf-8")
            with self.assertRaises(ValueError):
                fill_template(Path(tmp) / "db.sqlite", template, Path(tmp) / "out")


if __name__ == "__main__":
    unittest.main()

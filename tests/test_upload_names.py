import datetime as dt
import tempfile
import unittest
from pathlib import Path

from chint_etm_mdm.upload_names import report_path_for, series_label, series_slug, upload_csv_path

DAY = dt.date(2026, 10, 21)


class UploadNameTests(unittest.TestCase):
    def test_series_slug_is_latin_lowercase(self) -> None:
        self.assertEqual("nb1l-20", series_slug("NB1L-20"))
        self.assertEqual("nvf-acc", series_slug("NVF Аксессуары"))
        self.assertEqual("", series_slug("Без серии"))

    def test_label_lists_up_to_three_series(self) -> None:
        self.assertEqual("nxm_ns2", series_label(["NXM", "NS2", "NXM"]))
        self.assertEqual("ns2_nxm", series_label(["NXM", "NS2", "NS2"]))
        self.assertEqual("mix", series_label(["NP2", "NP8", "ND16", "NC6"]))
        self.assertEqual("500820", series_label(["", "Без серии"], "500820"))
        self.assertEqual("chint", series_label([]))

    def test_name_gets_a_version_when_taken(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            first = upload_csv_path(folder, "char", ["NDK"], when=DAY)
            self.assertEqual("upload_char_ndk_2026-10-21.csv", first.name)
            first.write_text("x")
            second = upload_csv_path(folder, "char", ["NDK"], when=DAY)
            self.assertEqual("upload_char_ndk_2026-10-21_v2.csv", second.name)
            self.assertEqual("upload_char_ndk_2026-10-21_v2_report.xlsx", report_path_for(second).name)
            with self.assertRaises(ValueError):
                upload_csv_path(folder, "other", ["NDK"], when=DAY)


if __name__ == "__main__":
    unittest.main()

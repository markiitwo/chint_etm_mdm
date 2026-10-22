import datetime as dt
import tempfile
import unittest
import urllib.error
from pathlib import Path

from chint_etm_mdm.site_info import (
    FOUND,
    MISMATCH,
    NOT_FOUND,
    SiteCache,
    parse_product,
    parse_search,
    site_lines,
    update_site_cache,
)

FIXTURES = Path(__file__).parent / "fixtures"
SEARCH = (FIXTURES / "ensmas_search_616173.html").read_text(encoding="utf-8")
PRODUCT = (FIXTURES / "ensmas_product_616173.html").read_text(encoding="utf-8")
PAGE_URL = "https://ensmas.ru/catalog/oborudovanie_nizkogo_napryazheniya/preobrazovateli_chastoty/nvf7/preobrazovatel_chastoty_nvf7_3_0t_s6/"


class FakeSite:
    def __init__(self, pages: dict[str, str]) -> None:
        self.pages = pages
        self.asked: list[str] = []

    def __call__(self, url: str) -> str:
        self.asked.append(url)
        if url not in self.pages:
            raise urllib.error.URLError("нет связи")
        return self.pages[url]


class SiteInfoTests(unittest.TestCase):
    def test_search_cards(self) -> None:
        cards = parse_search(SEARCH)
        self.assertEqual(1, len(cards))
        self.assertEqual("616173", cards[0].sku)
        self.assertEqual(PAGE_URL, cards[0].url)
        self.assertTrue(cards[0].image.startswith("https://ensmas.ru/upload/iblock/90b/"))

    def test_product_page(self) -> None:
        record = parse_product(PRODUCT, "616173")
        self.assertEqual(FOUND, record.status)
        self.assertEqual(PAGE_URL, record.url)
        # The first gallery picture, as the old parser took it, not og:image.
        self.assertEqual("https://ensmas.ru/upload/iblock/fd6/po8sbroq/%D0%A4%D0%BE%D1%82%D0%BE%20T7%20LED.jpg", record.image_url)
        self.assertEqual("ЕАЭС N RU Д-CN.РА06.В.41573/24", record.cert_number)
        self.assertEqual("https://ensmas.ru/upload/iblock/674/gqhgfk2a/Декларация о соответствии ЕАС.pdf", record.cert_pdf_url)
        self.assertEqual(
            [("Сертификаты", "Декларация о соответствии ЕАС"), ("Документы", "NVF7 - Паспорт устройства"), ("Документы", "3D NVF7")],
            [(doc["section"], doc["title"]) for doc in record.documents],
        )
        self.assertEqual("270х466х268.1", record.properties["Габариты"])
        self.assertEqual("IP20", record.properties["Степень защиты IP"])
        self.assertEqual("NVF7", record.properties["Серия"])

    def test_page_of_another_article_is_not_used(self) -> None:
        record = parse_product(PRODUCT, "616172")
        self.assertEqual(MISMATCH, record.status)
        self.assertIn("616173", record.note)

    def test_cache_and_lookup(self) -> None:
        site = FakeSite(
            {
                "https://ensmas.ru/catalog/?q=616173": SEARCH,
                PAGE_URL: PRODUCT,
                "https://ensmas.ru/catalog/?q=616172": SEARCH,  # search shows only the neighbour
            }
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "cache" / "ensmas_site.json"
            cache = SiteCache(path)
            result = update_site_cache(cache, ["616173", "616172", 616173.0], fetch=site, delay=0)
            self.assertEqual((2, 1, 1), (result.asked, result.found, result.not_found))
            self.assertEqual(FOUND, cache.records["616173"].status)
            self.assertEqual(NOT_FOUND, cache.records["616172"].status)
            self.assertIn("616173", cache.records["616172"].note)

            again = SiteCache(path)
            self.assertEqual("ЕАЭС N RU Д-CN.РА06.В.41573/24", again.found("616173").cert_number)
            self.assertIsNone(again.found("616172"))
            asked_before = len(site.asked)
            second = update_site_cache(again, ["616173", "616172"], fetch=site, delay=0)
            self.assertEqual((0, 2), (second.asked, second.from_cache))
            self.assertEqual(asked_before, len(site.asked))

            later = update_site_cache(again, ["616172"], fetch=site, delay=0, today=dt.date.today() + dt.timedelta(days=8))
            self.assertEqual(1, later.asked)  # "not found" is asked again after a week

    def test_site_down(self) -> None:
        site = FakeSite({})
        with tempfile.TemporaryDirectory() as tmp:
            cache = SiteCache(Path(tmp) / "site.json")
            lines: list[str] = []
            result = update_site_cache(cache, [str(n) for n in range(100, 120)], fetch=site, delay=0, progress=lines.append)
        self.assertEqual(5, len(result.errors))  # gives up instead of trying all twenty
        self.assertEqual({}, cache.records)  # errors are not remembered, asked next time
        self.assertIn("дальше без него", lines[-1])
        self.assertTrue(any("Ошибок связи" in line for line in site_lines(result)))

    def test_stop(self) -> None:
        site = FakeSite({"https://ensmas.ru/catalog/?q=1": "<html></html>"})
        cache = SiteCache(None)
        result = update_site_cache(cache, ["1", "2"], fetch=site, delay=0, should_stop=lambda: True)
        self.assertTrue(result.stopped)
        self.assertEqual(0, result.asked)


if __name__ == "__main__":
    unittest.main()

"""Product data from the CHINT site ensmas.ru: card link, picture, certificate.

The site search returns the product card for an article; the program opens the
card only when the article written on it is the one asked for (the old image
parser took the first card found and could pick another product). From the
product page it keeps:

- the page link (RAEK "Ссылка на сайте поставщика");
- the main picture (RAEK and ETM "Изображение"), the same file the old parser took;
- the certificate number and the PDF of the certificate;
- links to the other documents (passport, manual, catalog, 2D/3D);
- the characteristics the page lists, for later use.

Results are kept in a cache file in the work folder, so the site is asked only
about articles it was not asked about recently. A pause between requests keeps
the load on the site low.
"""
from __future__ import annotations

import datetime as dt
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable, Iterable
from dataclasses import asdict, dataclass, field
from html.parser import HTMLParser
from pathlib import Path

from .atomic_files import atomic_write_json
from .filler import normalize_article
from .net import open_url

BASE_URL = "https://ensmas.ru"
SEARCH_URL = BASE_URL + "/catalog/?q={}"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
REQUEST_TIMEOUT = 20
REQUEST_DELAY = 1.5  # seconds between requests, as in the old parser
FOUND, NOT_FOUND, MISMATCH = "found", "not_found", "mismatch"
REFRESH_DAYS = {FOUND: 30, NOT_FOUND: 7, MISMATCH: 7}
CERTIFICATE_PROPERTY = "номер сертификата"
CERTIFICATE_SECTION = "сертификат"

Fetch = Callable[[str], str]


@dataclass(frozen=True)
class SiteCard:
    sku: str
    url: str
    image: str


@dataclass
class SiteRecord:
    article: str
    status: str
    url: str = ""
    image_url: str = ""
    name: str = ""
    cert_number: str = ""
    cert_pdf_url: str = ""
    documents: list[dict[str, str]] = field(default_factory=list)
    properties: dict[str, str] = field(default_factory=dict)
    note: str = ""
    checked_at: str = ""


@dataclass
class SiteResult:
    asked: int = 0
    found: int = 0
    not_found: int = 0
    errors: list[str] = field(default_factory=list)
    from_cache: int = 0
    stopped: bool = False


def absolute(url: str) -> str:
    return urllib.parse.urljoin(BASE_URL + "/", url.strip()) if url else ""


def product_photo(url: str) -> bool:
    """A photo of the product itself. Pages without one carry the site logo in og:image
    (/upload/dev2fun_opengraph/); only pictures of the catalogue (/upload/iblock/) count."""
    return "/upload/iblock/" in urllib.parse.unquote(urllib.parse.urlsplit(url).path)


def classes(attrs: dict[str, str | None]) -> set[str]:
    return set((attrs.get("class") or "").split())


def clean(text: str) -> str:
    return " ".join(text.split())


class SearchParser(HTMLParser):
    """Product cards of a catalog search page (schema.org microdata)."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.cards: list[dict[str, str]] = []

    def handle_starttag(self, tag: str, attr_list: list[tuple[str, str | None]]) -> None:
        attrs = dict(attr_list)
        if tag == "div" and "product-card" in classes(attrs):
            self.cards.append({})
            return
        if not self.cards:
            return
        card = self.cards[-1]
        prop = attrs.get("itemprop")
        if tag == "meta" and prop == "sku":
            card.setdefault("sku", attrs.get("content") or "")
        elif tag == "a" and prop == "url":
            card.setdefault("url", attrs.get("href") or "")
        elif tag == "img" and prop == "image":
            card.setdefault("image", attrs.get("src") or "")


def parse_search(html: str) -> list[SiteCard]:
    parser = SearchParser()
    parser.feed(html)
    return [
        SiteCard(normalize_article(card.get("sku", "")), absolute(card.get("url", "")), absolute(card.get("image", "")))
        for card in parser.cards
        if card.get("url")
    ]


class ProductParser(HTMLParser):
    """The parts of a product page the program keeps."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.json_ld: list[str] = []
        self.canonical = ""
        self.og_url = ""
        self.og_image = ""
        self.gallery: list[str] = []
        self.articul = ""
        self.properties: dict[str, str] = {}
        self.documents: list[dict[str, str]] = []
        self._capture: str | None = None  # what the text being read belongs to
        self._capture_tag = ""
        self._capture_depth = 0
        self._buffer: list[str] = []
        self._section = ""
        self._title = ""
        self._doc: dict[str, str] | None = None
        self._doc_depth = 0

    def handle_starttag(self, tag: str, attr_list: list[tuple[str, str | None]]) -> None:
        attrs = dict(attr_list)
        names = classes(attrs)
        if self._capture is not None and tag == self._capture_tag:
            self._capture_depth += 1
        if tag == "link" and (attrs.get("rel") or "").lower() == "canonical":
            self.canonical = attrs.get("href") or ""
        elif tag == "meta" and attrs.get("property") == "og:url":
            self.og_url = attrs.get("content") or ""
        elif tag == "meta" and attrs.get("property") == "og:image":
            self.og_image = attrs.get("content") or ""
        elif tag == "a" and attrs.get("data-fancybox") == "gal" and attrs.get("href"):
            self.gallery.append(attrs["href"] or "")
        elif tag == "a" and "link-with-icon" in names and attrs.get("href"):
            self._doc = {"section": self._section, "title": "", "url": absolute(attrs["href"] or "")}
            self._doc_depth = 1
            return
        elif self._doc is not None and tag == "a":
            self._doc_depth += 1
        if self._capture is not None:
            return
        if tag == "script" and (attrs.get("type") or "").lower() == "application/ld+json":
            self._start("json_ld", tag)
        elif "articul" in names:
            self._start("articul", tag)
        elif tag == "p" and "documents-wrapper-title" in names:
            self._start("section", tag)
        elif tag == "div" and names == {"title"}:
            self._start("prop_title", tag)
        elif tag == "div" and names == {"value"} and self._title:
            self._start("prop_value", tag)

    def handle_endtag(self, tag: str) -> None:
        if self._doc is not None and tag == "a":
            self._doc_depth -= 1
            if self._doc_depth == 0:
                self._doc["title"] = clean(self._doc["title"]) or Path(urllib.parse.unquote(self._doc["url"])).stem
                self.documents.append(self._doc)
                self._doc = None
            return
        if self._capture is None or tag != self._capture_tag:
            return
        self._capture_depth -= 1
        if self._capture_depth:
            return
        text = "".join(self._buffer)
        kind, self._capture, self._buffer = self._capture, None, []
        if kind == "json_ld":
            self.json_ld.append(text)
        elif kind == "articul":
            self.articul = clean(text)
        elif kind == "section":
            self._section = clean(text)
        elif kind == "prop_title":
            self._title = clean(text)
        elif kind == "prop_value":
            if self._title:
                self.properties.setdefault(self._title, clean(text))
            self._title = ""

    def handle_data(self, data: str) -> None:
        if self._doc is not None:
            self._doc["title"] += data
        if self._capture is not None:
            self._buffer.append(data)

    def _start(self, kind: str, tag: str) -> None:
        self._capture = kind
        self._capture_tag = tag
        self._capture_depth = 1
        self._buffer = []


def json_ld_product(blocks: Iterable[str]) -> dict:
    for block in blocks:
        try:
            data = json.loads(block)
        except ValueError:
            continue
        items = data if isinstance(data, list) else data.get("@graph", [data]) if isinstance(data, dict) else []
        for item in items:
            if isinstance(item, dict) and str(item.get("@type", "")).lower() == "product":
                return item
    return {}


def first_text(value: object) -> str:
    if isinstance(value, list):
        value = value[0] if value else ""
    if isinstance(value, dict):
        value = value.get("url") or value.get("contentUrl") or ""
    return str(value or "").strip()


def parse_product(html: str, article: str) -> SiteRecord:
    """A product page. The status is MISMATCH when the page is about another article."""
    parser = ProductParser()
    parser.feed(html)
    product = json_ld_product(parser.json_ld)
    properties = dict(parser.properties)
    for item in product.get("additionalProperty") or []:
        if isinstance(item, dict) and item.get("name"):
            properties.setdefault(clean(str(item["name"])), clean(str(item.get("value", ""))))
    sku = normalize_article(first_text(product.get("sku")) or re.sub(r"(?i)^\s*артикул\s*:?\s*", "", parser.articul))
    url = absolute(parser.og_url or parser.canonical or first_text(product.get("url")))
    candidates = [*parser.gallery, first_text(product.get("image")), parser.og_image]
    image = next((url for url in map(absolute, candidates) if url and product_photo(url)), "")
    cert_number = next((value for name, value in properties.items() if name.casefold() == CERTIFICATE_PROPERTY), "")
    cert_pdf = next(
        (doc["url"] for doc in parser.documents if CERTIFICATE_SECTION in doc["section"].casefold() and doc["url"].lower().endswith(".pdf")),
        "",
    )
    record = SiteRecord(
        article=article,
        status=FOUND,
        url=url,
        image_url=image,
        name=clean(first_text(product.get("name"))),
        cert_number=cert_number,
        cert_pdf_url=cert_pdf,
        documents=parser.documents,
        properties=properties,
    )
    if sku != article:
        record.status = MISMATCH
        record.note = f"На странице другой артикул: {sku or 'не указан'}"
    return record


def fetch_url(url: str) -> str:
    """GET a page as text. The program's own certificate bundle is used when the system one is missing."""
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "ru"})
    response = open_url(request, REQUEST_TIMEOUT)
    with response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def look_up(article: str, fetch: Fetch, pause: Callable[[], None]) -> SiteRecord:
    cards = parse_search(fetch(SEARCH_URL.format(urllib.parse.quote(article))))
    card = next((card for card in cards if card.sku == article), None)
    if card is None:
        others = ", ".join(card.sku for card in cards[:3] if card.sku)
        note = f"Поиск нашёл другие артикулы: {others}" if others else "Сайт не нашёл товар с таким артикулом"
        return SiteRecord(article, NOT_FOUND, note=note)
    pause()
    record = parse_product(fetch(card.url), article)
    record.url = record.url or card.url
    if not record.image_url and product_photo(card.image):
        record.image_url = card.image
    return record


class SiteCache:
    """Answers of the site by article, in a JSON file of the work folder."""

    def __init__(self, path: Path | None) -> None:
        self.path = path
        self.records: dict[str, SiteRecord] = {}
        if path and path.exists():
            data = json.loads(path.read_text(encoding="utf-8"))
            for article, item in data.get("products", {}).items():
                record = SiteRecord(**item)
                if record.image_url and not product_photo(record.image_url):
                    record.image_url = ""
                self.records[article] = record

    def fresh(self, article: str, today: dt.date) -> bool:
        record = self.records.get(article)
        if record is None or not record.checked_at:
            return False
        checked = dt.date.fromisoformat(record.checked_at[:10])
        return (today - checked).days < REFRESH_DAYS.get(record.status, 0)

    def save(self) -> None:
        if self.path is None:
            return
        self.path.parent.mkdir(parents=True, exist_ok=True)
        atomic_write_json(
            self.path,
            {"version": 1, "source": BASE_URL, "products": {article: asdict(record) for article, record in sorted(self.records.items())}},
        )

    def found(self, article: str) -> SiteRecord | None:
        record = self.records.get(article)
        return record if record and record.status == FOUND else None


def update_site_cache(
    cache: SiteCache,
    articles: Iterable[str],
    fetch: Fetch = fetch_url,
    delay: float = REQUEST_DELAY,
    progress: Callable[[str], None] | None = None,
    should_stop: Callable[[], bool] | None = None,
    today: dt.date | None = None,
) -> SiteResult:
    """Ask the site about articles the cache has no recent answer for; save after every few products."""
    today = today or dt.date.today()
    todo = []
    result = SiteResult()
    for article in dict.fromkeys(normalize_article(item) for item in articles):
        if not article:
            continue
        if cache.fresh(article, today):
            result.from_cache += 1
        else:
            todo.append(article)
    pause = (lambda: time.sleep(delay)) if delay > 0 else (lambda: None)
    for number, article in enumerate(todo, 1):
        if should_stop and should_stop():
            result.stopped = True
            break
        if number > 1:
            pause()
        result.asked += 1
        try:
            record = look_up(article, fetch, pause)
        except (urllib.error.URLError, OSError, ValueError) as exc:
            result.errors.append(f"{article}: {exc}")
            if progress:
                progress(f"Сайт: {article} — ошибка связи ({exc})")
            if len(result.errors) >= 5 and result.found == 0 and result.not_found == 0:
                if progress:
                    progress("Сайт не отвечает, дальше без него.")
                break
            continue
        record.checked_at = dt.datetime.now().isoformat(timespec="seconds")
        cache.records[article] = record
        if record.status == FOUND:
            result.found += 1
        else:
            result.not_found += 1
        if progress:
            progress(f"Сайт: {number} из {len(todo)} — {article} {'найден' if record.status == FOUND else 'не найден'}")
        if number % 10 == 0:
            cache.save()
    cache.save()
    return result


def default_cache_path(work_dir: Path) -> Path:
    return work_dir / "cache" / "ensmas_site.json"


def site_lines(result: SiteResult) -> list[str]:
    lines = []
    if result.asked or result.from_cache:
        lines.append(
            f"Сайт ensmas.ru: спросил {result.asked}, найдено {result.found}, не найдено {result.not_found}; "
            f"уже было в памяти {result.from_cache}."
        )
    if result.errors:
        lines.append(f"Ошибок связи с сайтом: {len(result.errors)}. Эти товары программа спросит в следующий раз.")
    if result.stopped:
        lines.append("Проверка сайта остановлена.")
    return lines

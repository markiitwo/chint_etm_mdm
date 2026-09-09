from __future__ import annotations

import datetime as dt
import html.parser
import os
import posixpath
import re
import sqlite3
import tempfile
import urllib.parse
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .db_utils import connect_database, create_sqlite_backup


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
PRICE_SHEETS = (
    "Price-list",
    "Введено в ассортимент",
    "Выведено из ассортимента",
    "История изменений",
)
REQUIRED_TABLE_HEADERS = ("Артикул", "Наименование")
DEFAULT_PRICE_SOURCE_URL = "https://ensmas.ru/"
LOW_VOLTAGE_PRICE_LABEL = "низковольтное оборудование"
MAX_PRICE_DOWNLOAD_BYTES = 100 * 1024 * 1024


@dataclass(frozen=True)
class PriceImportResult:
    price_file: Path
    db_path: Path
    backup_path: Path | None
    snapshot_id: int
    snapshot_date: str
    price_rows: int
    new_rows: int
    out_rows: int
    history_rows: int
    products_total: int
    new_articles: int
    changed_price_rows: int


class PriceLinkParser(html.parser.HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.links: list[tuple[str, str]] = []
        self._current_href: str | None = None
        self._current_text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        attrs_dict = {name.lower(): value or "" for name, value in attrs}
        self._current_href = attrs_dict.get("href")
        self._current_text = []

    def handle_data(self, data: str) -> None:
        if self._current_href is not None:
            self._current_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() != "a" or self._current_href is None:
            return
        text = " ".join("".join(self._current_text).split())
        self.links.append((self._current_href, text))
        self._current_href = None
        self._current_text = []


def find_latest_price_url(source_url: str = DEFAULT_PRICE_SOURCE_URL) -> str:
    source_url = source_url.strip() or DEFAULT_PRICE_SOURCE_URL
    request = urllib.request.Request(source_url, headers={"User-Agent": "CHINT ETM MDM"})
    with urllib.request.urlopen(request, timeout=30) as response:
        raw_html = response.read()
        encoding = response.headers.get_content_charset() or "utf-8"
    page_html = raw_html.decode(encoding, errors="replace")

    parser = PriceLinkParser()
    parser.feed(page_html)

    price_links: list[str] = []
    for href, text in parser.links:
        normalized_text = text.casefold()
        if LOW_VOLTAGE_PRICE_LABEL in normalized_text and "Price-list-CHINT" in href:
            return urllib.parse.urljoin(source_url, href)
        if "Price-list-CHINT" in href and href.lower().endswith((".xlsx", ".xlsm")):
            price_links.append(href)
        elif "Price-list-CHINT" in href and ".xlsx" in href.lower():
            price_links.append(href)

    if price_links:
        return urllib.parse.urljoin(source_url, price_links[0])
    raise ValueError("Не удалось найти ссылку на прайс-лист CHINT на сайте.")


def download_price_file(url: str, downloads_dir: Path) -> Path:
    url = url.strip()
    if not url:
        raise ValueError("Укажите ссылку на прайс-лист.")
    downloads_dir.mkdir(parents=True, exist_ok=True)
    request = urllib.request.Request(url, headers={"User-Agent": "CHINT ETM MDM"})
    temp_path: Path | None = None
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            filename = filename_from_response(url, response.headers.get("Content-Disposition", ""))
            content_length_text = response.headers.get("Content-Length", "")
            try:
                content_length = int(content_length_text) if content_length_text else None
            except ValueError:
                content_length = None
            if content_length is not None and content_length > MAX_PRICE_DOWNLOAD_BYTES:
                raise ValueError("Скачиваемый прайс превышает допустимый размер.")
            with tempfile.NamedTemporaryFile(
                mode="wb", prefix=".price-", suffix=".part", dir=downloads_dir, delete=False
            ) as fh:
                temp_path = Path(fh.name)
                downloaded = 0
                while chunk := response.read(1024 * 1024):
                    downloaded += len(chunk)
                    if downloaded > MAX_PRICE_DOWNLOAD_BYTES:
                        raise ValueError("Скачиваемый прайс превышает допустимый размер.")
                    fh.write(chunk)

        if content_length is not None and downloaded != content_length:
            raise ValueError("Скачанный файл короче размера, указанного сервером.")

        validate_excel_package(temp_path)
        target = publish_unique_path(temp_path, downloads_dir / filename)
        temp_path = None
        return target
    finally:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)


def filename_from_response(url: str, content_disposition: str) -> str:
    extended = re.search(r"(?:^|;)\s*filename\*\s*=\s*(?:UTF-8'[^']*')?([^;]+)", content_disposition, re.I)
    regular = re.search(r'(?:^|;)\s*filename\s*=\s*(?:"([^"]*)"|([^;]*))', content_disposition, re.I)
    if extended:
        name = urllib.parse.unquote(extended.group(1).strip().strip('"'))
    elif regular:
        name = (regular.group(1) if regular.group(1) is not None else regular.group(2)).strip()
    else:
        name = urllib.parse.unquote(urllib.parse.urlparse(url).path.rsplit("/", 1)[-1])
    name = re.split(r"[/\\]", name)[-1]
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).rstrip(" .")
    if not name:
        stamp = dt.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        name = f"Price-list-CHINT_{stamp}.xlsx"
    if not name.lower().endswith((".xlsx", ".xlsm")):
        name = f"{name}.xlsx"
    if name.split(".", 1)[0].upper() in {
        "CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))
    }:
        name = f"_{name}"
    return name


def validate_excel_package(path: Path) -> None:
    if path.stat().st_size == 0:
        raise ValueError("Скачанный файл пуст.")
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            required = {"[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels"}
            if not required.issubset(names):
                raise ValueError("Скачанный файл не содержит структуру книги Excel.")
            if archive.testzip() is not None:
                raise ValueError("Скачанный ZIP-архив поврежден.")
            ET.fromstring(archive.read("[Content_Types].xml"))
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            if workbook.tag != f"{{{NS['a']}}}workbook" or workbook.find("a:sheets", NS) is None:
                raise ValueError("Скачанный файл не содержит корректную книгу Excel.")
            worksheet_targets = [
                rel.attrib.get("Target", "")
                for rel in rels
                if rel.attrib.get("Type", "").endswith("/worksheet")
                and rel.attrib.get("TargetMode") != "External"
            ]
            if not worksheet_targets:
                raise ValueError("Скачанная книга Excel не содержит листов.")
            for target in worksheet_targets:
                member = posixpath.normpath(
                    target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
                )
                if member.startswith("../") or member not in names:
                    raise ValueError("Скачанная книга Excel ссылается на отсутствующий лист.")
    except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise ValueError("Скачанный файл не является корректной книгой Excel.") from exc


def publish_unique_path(temp_path: Path, path: Path) -> Path:
    for index in range(1000):
        candidate = path if index == 0 else path.with_name(f"{path.stem}_{index}{path.suffix}")
        try:
            if os.name == "nt":
                os.rename(temp_path, candidate)
            else:
                os.link(temp_path, candidate)
        except FileExistsError:
            continue
        if os.name != "nt":
            temp_path.unlink()
        return candidate
    raise ValueError(f"Не удалось подобрать свободное имя файла для {path}")


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(1, 1000):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise ValueError(f"Не удалось подобрать свободное имя файла для {path}")


def backup_database(db_path: Path) -> Path:
    return create_sqlite_backup(db_path).path


def cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    value = cell.find("a:v", NS)
    inline = cell.find("a:is", NS)
    if cell_type == "s" and value is not None and value.text:
        return shared_strings[int(value.text)]
    if cell_type == "inlineStr" and inline is not None:
        return "".join(item.text or "" for item in inline.iterfind(".//a:t", NS))
    return value.text if value is not None and value.text is not None else ""


def col_to_num(col: str) -> int:
    number = 0
    for char in col:
        if char.isalpha():
            number = number * 26 + (ord(char.upper()) - 64)
    return number


def num_to_col(number: int) -> str:
    result = ""
    while number:
        number, rem = divmod(number - 1, 26)
        result = chr(65 + rem) + result
    return result


def read_sheet_rows(xlsx_path: Path, sheet_name: str) -> list[list[str]]:
    with zipfile.ZipFile(xlsx_path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel for rel in rels}
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", NS):
                shared_strings.append("".join(t.text or "" for t in item.iterfind(".//a:t", NS)))

        target_sheet = None
        sheets = workbook.find("a:sheets", NS)
        for sheet in sheets if sheets is not None else []:
            if sheet.attrib["name"] == sheet_name:
                target_sheet = sheet
                break
        if target_sheet is None:
            raise ValueError(f"В прайсе нет листа '{sheet_name}'.")

        rid = target_sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        relation = relmap.get(rid or "")
        if relation is None:
            raise ValueError(f"Для листа '{sheet_name}' не найдена связь в книге.")
        if not relation.attrib.get("Type", "").endswith("/worksheet") or relation.attrib.get("TargetMode") == "External":
            raise ValueError(f"Связь листа '{sheet_name}' не указывает на лист внутри книги.")
        target = relation.attrib.get("Target", "").replace("\\", "/")
        sheet_path = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target))
        if sheet_path.startswith("../") or sheet_path not in archive.namelist():
            raise ValueError(f"Файл листа '{sheet_name}' отсутствует в книге.")
        root = ET.fromstring(archive.read(sheet_path))
        rows_xml = root.findall(".//a:sheetData/a:row", NS)
        max_col = 0
        sparse_rows: list[dict[str, str]] = []
        for row in rows_xml:
            sparse: dict[str, str] = {}
            for cell in row.findall("a:c", NS):
                ref = cell.attrib["r"]
                col = "".join(char for char in ref if char.isalpha())
                sparse[col] = cell_value(cell, shared_strings)
                max_col = max(max_col, col_to_num(col))
            sparse_rows.append(sparse)
        return [
            [sparse.get(num_to_col(index), "") for index in range(1, max_col + 1)]
            for sparse in sparse_rows
        ]


def table_rows_from_header(rows: list[list[str]], sheet_name: str) -> list[list[str]]:
    for header_index, row in enumerate(rows):
        header = [str(cell).strip() for cell in row]
        if all(required in header for required in REQUIRED_TABLE_HEADERS):
            return [header, *rows[header_index + 1 :]]
    required = ", ".join(REQUIRED_TABLE_HEADERS)
    raise ValueError(f"На листе '{sheet_name}' не найдены обязательные заголовки: {required}.")


def read_price_workbook(xlsx_path: Path) -> dict[str, list[list[str]]]:
    if xlsx_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Прайс должен быть XLSX/XLSM файлом.")
    workbook_rows = {
        sheet_name: table_rows_from_header(read_sheet_rows(xlsx_path, sheet_name), sheet_name)
        for sheet_name in PRICE_SHEETS
    }
    price_rows = workbook_rows["Price-list"]
    article_index = price_rows[0].index("Артикул")
    if not any(
        article_index < len(row) and str(row[article_index]).strip()
        for row in price_rows[1:]
    ):
        raise ValueError("Лист 'Price-list' не содержит строк с артикулами.")
    return workbook_rows


def snapshot_date_from_filename(path: Path) -> str:
    return path.stem.replace("Price-list-CHINT_", "")


def latest_snapshot_id(conn: sqlite3.Connection) -> int | None:
    row = conn.execute(
        """
        SELECT snapshots.id
        FROM price_snapshots snapshots
        WHERE EXISTS (
            SELECT 1
            FROM price_snapshot_items items
            WHERE items.snapshot_id = snapshots.id
        )
        ORDER BY snapshots.id DESC
        LIMIT 1
        """
    ).fetchone()
    return int(row[0]) if row else None


def count_new_articles(conn: sqlite3.Connection, previous_snapshot_id: int | None, snapshot_id: int) -> int:
    if previous_snapshot_id is None:
        return int(
            conn.execute(
                "SELECT COUNT(DISTINCT article) FROM price_snapshot_items WHERE snapshot_id = ?",
                (snapshot_id,),
            ).fetchone()[0]
            or 0
        )
    return int(
        conn.execute(
            """
            SELECT COUNT(DISTINCT current.article)
            FROM price_snapshot_items current
            WHERE current.snapshot_id = ?
              AND NOT EXISTS (
                  SELECT 1
                  FROM price_snapshot_items previous
                  WHERE previous.snapshot_id = ?
                    AND previous.article = current.article
              )
            """,
            (snapshot_id, previous_snapshot_id),
        ).fetchone()[0]
        or 0
    )


def count_changed_prices(conn: sqlite3.Connection, previous_snapshot_id: int | None, snapshot_id: int) -> int:
    if previous_snapshot_id is None:
        return 0
    return int(
        conn.execute(
            """
            SELECT COUNT(*)
            FROM price_snapshot_items current
            JOIN price_snapshot_items previous ON previous.article = current.article
            WHERE current.snapshot_id = ?
              AND previous.snapshot_id = ?
              AND (
                  COALESCE(current.price_with_vat, '') != COALESCE(previous.price_with_vat, '')
                  OR COALESCE(current.price_without_vat, '') != COALESCE(previous.price_without_vat, '')
              )
            """,
            (snapshot_id, previous_snapshot_id),
        ).fetchone()[0]
        or 0
    )


def row_dict(header: list[str], row: list[str]) -> dict[str, str]:
    return dict(zip(header, row))


def value(data: dict[str, str], name: str) -> str:
    return data.get(name, "")


def insert_price_rows(cur: sqlite3.Cursor, snapshot_id: int, rows: list[list[str]]) -> int:
    if not rows:
        return 0
    header = rows[0]
    imported = 0
    for row in rows[1:]:
        if not any(row):
            continue
        data = row_dict(header, row)
        article = value(data, "Артикул").strip()
        if not article:
            continue
        imported += 1
        cur.execute(
            """
            INSERT INTO products (
                article, name, series_name, equipment_type, equipment_group,
                price_collection, unit, stock_status, is_new, is_promo, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(article) DO UPDATE SET
                name=excluded.name,
                series_name=excluded.series_name,
                equipment_type=excluded.equipment_type,
                equipment_group=excluded.equipment_group,
                price_collection=excluded.price_collection,
                unit=excluded.unit,
                stock_status=excluded.stock_status,
                is_new=excluded.is_new,
                is_promo=excluded.is_promo,
                updated_at=CURRENT_TIMESTAMP
            """,
            (
                article,
                value(data, "Наименование"),
                value(data, "Серия"),
                value(data, "Тип оборудования"),
                value(data, "Группа оборудования"),
                value(data, "Коллекция"),
                value(data, "Ед."),
                value(data, "Складской статус"),
                value(data, "Новинка"),
                value(data, "Акция"),
            ),
        )
        cur.execute(
            """
            INSERT INTO price_snapshot_items (
                snapshot_id, article, name, price_with_vat, price_without_vat, unit,
                price_collection, stock_status, is_removed, is_new, is_promo,
                equipment_type, equipment_group, series_name, pack_multiple,
                min_shipment, gross_weight_transport, net_weight_transport,
                transport_volume, gross_weight_min, net_weight_min, min_volume,
                gross_weight_unit, net_weight_unit, unit_volume, outgoing_stock,
                analog_article, analog_name
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                snapshot_id,
                article,
                value(data, "Наименование"),
                value(data, "Тариф с НДС, руб"),
                value(data, "Тариф без НДС, руб"),
                value(data, "Ед."),
                value(data, "Коллекция"),
                value(data, "Складской статус"),
                value(data, "Вывод из ассортимента"),
                value(data, "Новинка"),
                value(data, "Акция"),
                value(data, "Тип оборудования"),
                value(data, "Группа оборудования"),
                value(data, "Серия"),
                value(data, "Кратность транспортной упаковки"),
                value(data, "Минимальная отгрузка"),
                value(data, "Вес брутто транспортной упаковки, кг"),
                value(data, "Вес нетто транспортной упаковки, кг"),
                value(data, "Объём транспортной упаковки, м^3"),
                value(data, "Вес брутто минимальной отгрузки, кг"),
                value(data, "Вес нетто минимальной отгрузки, кг"),
                value(data, "Объем минимальной отгрузки, м^3"),
                value(data, "Вес брутто за 1 ед. прод."),
                value(data, "Вес  нетто за 1 ед. прод."),
                value(data, "Объём, за 1 ед. прод."),
                value(data, "Остатки выводимой продукции"),
                value(data, "Артикул аналога"),
                value(data, "Наименование аналога"),
            ),
        )
    return imported


def insert_new_products(cur: sqlite3.Cursor, snapshot_id: int, rows: list[list[str]]) -> int:
    if not rows:
        return 0
    header = rows[0]
    imported = 0
    for row in rows[1:]:
        if not any(row):
            continue
        data = row_dict(header, row)
        article = value(data, "Артикул").strip()
        if not article:
            continue
        imported += 1
        cur.execute(
            """
            INSERT INTO new_products (article, name, introduced_at, first_snapshot_id)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(article) DO UPDATE SET
                name=excluded.name,
                introduced_at=excluded.introduced_at
            """,
            (article, value(data, "Наименование"), value(data, "Дата ввода"), snapshot_id),
        )
    return imported


def insert_outgoing(cur: sqlite3.Cursor, snapshot_id: int, rows: list[list[str]]) -> int:
    if not rows:
        return 0
    header = rows[0]
    imported = 0
    cur.execute("DELETE FROM assortment_outgoing WHERE snapshot_id = ?", (snapshot_id,))
    for row in rows[1:]:
        if not any(row):
            continue
        data = row_dict(header, row)
        article = value(data, "Артикул").strip()
        if not article:
            continue
        imported += 1
        cur.execute(
            """
            INSERT INTO assortment_outgoing (article, name, removed_at, snapshot_id)
            VALUES (?, ?, ?, ?)
            """,
            (article, value(data, "Наименование"), value(data, "Дата вывода"), snapshot_id),
        )
    return imported


def insert_history(cur: sqlite3.Cursor, snapshot_id: int, rows: list[list[str]]) -> int:
    if not rows:
        return 0
    header = rows[0]
    imported = 0
    cur.execute("DELETE FROM price_change_history WHERE snapshot_id = ?", (snapshot_id,))
    for row in rows[1:]:
        if not any(row):
            continue
        data = row_dict(header, row)
        article = value(data, "Артикул").strip()
        if not article:
            continue
        imported += 1
        cur.execute(
            """
            INSERT INTO price_change_history (
                article, name, change_type, old_value, new_value, changed_at, snapshot_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                article,
                value(data, "Наименование"),
                value(data, "Что изменилось"),
                value(data, "Старое значение"),
                value(data, "Новое значение"),
                value(data, "Дата изменения"),
                snapshot_id,
            ),
        )
    return imported


def import_price_workbook(
    xlsx_path: Path,
    db_path: Path,
    source_url: str = "",
    make_backup: bool = True,
) -> PriceImportResult:
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Прайс не найден: {xlsx_path}")
    if not db_path.exists():
        raise FileNotFoundError(f"База не найдена: {db_path}")

    workbook_rows = read_price_workbook(xlsx_path)
    backup_path = backup_database(db_path) if make_backup else None
    snapshot_date = snapshot_date_from_filename(xlsx_path)

    with closing(connect_database(db_path)) as conn, conn:
        previous_snapshot_id = latest_snapshot_id(conn)
        cur = conn.cursor()
        cur.execute(
            """
            INSERT INTO price_snapshots (snapshot_date, source_url, local_file, workbook_name)
            VALUES (?, ?, ?, ?)
            """,
            (snapshot_date, source_url, str(xlsx_path), xlsx_path.name),
        )
        snapshot_id = int(cur.lastrowid)

        price_rows = insert_price_rows(cur, snapshot_id, workbook_rows["Price-list"])
        new_rows = insert_new_products(cur, snapshot_id, workbook_rows["Введено в ассортимент"])
        out_rows = insert_outgoing(cur, snapshot_id, workbook_rows["Выведено из ассортимента"])
        history_rows = insert_history(cur, snapshot_id, workbook_rows["История изменений"])
        conn.commit()

        products_total = int(cur.execute("SELECT COUNT(*) FROM products").fetchone()[0])
        new_articles = count_new_articles(conn, previous_snapshot_id, snapshot_id)
        changed_price_rows = count_changed_prices(conn, previous_snapshot_id, snapshot_id)

    return PriceImportResult(
        price_file=xlsx_path,
        db_path=db_path,
        backup_path=backup_path,
        snapshot_id=snapshot_id,
        snapshot_date=snapshot_date,
        price_rows=price_rows,
        new_rows=new_rows,
        out_rows=out_rows,
        history_rows=history_rows,
        products_total=products_total,
        new_articles=new_articles,
        changed_price_rows=changed_price_rows,
    )


def result_lines(result: PriceImportResult) -> list[str]:
    lines = [
        f"Прайс: {result.price_file}",
        f"База: {result.db_path}",
        f"Снимок: {result.snapshot_date} (id {result.snapshot_id})",
        f"Строк прайса: {result.price_rows}",
        f"Новых товаров в листе: {result.new_rows}",
        f"Выведено из ассортимента: {result.out_rows}",
        f"Строк истории изменений: {result.history_rows}",
        f"Новых артикулов относительно прошлого прайса: {result.new_articles}",
        f"Строк с изменением цены относительно прошлого прайса: {result.changed_price_rows}",
        f"Товаров в базе всего: {result.products_total}",
    ]
    if result.backup_path:
        lines.append(f"Бэкап базы: {result.backup_path}")
    return lines

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class DatabaseStats:
    products_count: int
    dimensions_count: int
    attributes_count: int
    latest_price_snapshot: str
    latest_price_file: str


@dataclass(frozen=True)
class ProductRecord:
    article: str
    name: str = ""
    full_name: str = ""
    class81_code: str = ""
    unit: str = ""
    length_mm: float | None = None
    width_mm: float | None = None
    height_mm: float | None = None
    weight_kg: float | None = None
    volume_m3: float | None = None
    tnved_code: str = ""
    okpd2_code: str = ""
    attributes: dict[str, str] | None = None


def connect(db_path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _scalar(conn: sqlite3.Connection, query: str, params: tuple[Any, ...] = ()) -> Any:
    row = conn.execute(query, params).fetchone()
    return row[0] if row else None


def get_stats(db_path: Path) -> DatabaseStats:
    with connect(db_path) as conn:
        latest = conn.execute(
            """
            SELECT COALESCE(snapshot_date, '') AS snapshot_date,
                   COALESCE(workbook_name, local_file, '') AS workbook_name
            FROM price_snapshots
            ORDER BY id DESC
            LIMIT 1
            """
        ).fetchone()
        return DatabaseStats(
            products_count=int(_scalar(conn, "SELECT COUNT(*) FROM products") or 0),
            dimensions_count=int(
                _scalar(conn, "SELECT COUNT(*) FROM product_dimensions_resolved") or 0
            ),
            attributes_count=int(
                _scalar(conn, "SELECT COUNT(*) FROM product_attribute_values") or 0
            ),
            latest_price_snapshot=latest["snapshot_date"] if latest else "",
            latest_price_file=latest["workbook_name"] if latest else "",
        )


def fetch_products(db_path: Path, articles: list[str]) -> dict[str, ProductRecord]:
    clean_articles = sorted({a.strip() for a in articles if a and a.strip()})
    if not clean_articles:
        return {}

    placeholders = ",".join("?" for _ in clean_articles)
    products: dict[str, ProductRecord] = {}
    with connect(db_path) as conn:
        rows = conn.execute(
            f"""
            SELECT
                p.article,
                COALESCE(ig.name, p.name, '') AS name,
                COALESCE(ig.full_name, ig.name, p.name, '') AS full_name,
                COALESCE(ig.class81_code, '') AS class81_code,
                COALESCE(p.unit, ig.manufacturer, '') AS unit,
                d.length_mm,
                d.width_mm,
                d.height_mm,
                d.weight_kg,
                d.volume_m3,
                COALESCE((
                    SELECT tnved_code
                    FROM product_tnved_codes t
                    WHERE t.article = p.article
                    ORDER BY t.import_id DESC, t.id DESC
                    LIMIT 1
                ), '') AS tnved_code,
                COALESCE((
                    SELECT COALESCE(okpd2_9, okpd2_6, '')
                    FROM product_okpd2_codes o
                    WHERE o.article = p.article
                    ORDER BY o.import_id DESC, o.id DESC
                    LIMIT 1
                ), '') AS okpd2_code
            FROM products p
            LEFT JOIN (
                SELECT *
                FROM ipro_goods
                WHERE id IN (
                    SELECT MAX(id)
                    FROM ipro_goods
                    GROUP BY article
                )
            ) ig ON ig.article = p.article
            LEFT JOIN product_dimensions_resolved d ON d.article = p.article
            WHERE p.article IN ({placeholders})
            """,
            clean_articles,
        ).fetchall()

        for row in rows:
            article = str(row["article"])
            unit = str(row["unit"] or "").strip() or "шт"
            products[article] = ProductRecord(
                article=article,
                name=str(row["name"] or ""),
                full_name=str(row["full_name"] or ""),
                class81_code=str(row["class81_code"] or ""),
                unit=unit,
                length_mm=row["length_mm"],
                width_mm=row["width_mm"],
                height_mm=row["height_mm"],
                weight_kg=row["weight_kg"],
                volume_m3=row["volume_m3"],
                tnved_code=str(row["tnved_code"] or ""),
                okpd2_code=str(row["okpd2_code"] or ""),
                attributes={},
            )

        attr_rows = conn.execute(
            f"""
            SELECT article, attribute_name, value_normalized, value_raw
            FROM product_attribute_values
            WHERE article IN ({placeholders})
              AND COALESCE(attribute_name, '') <> ''
            ORDER BY confidence DESC, id DESC
            """,
            clean_articles,
        ).fetchall()

    for row in attr_rows:
        article = str(row["article"])
        product = products.get(article)
        if not product or product.attributes is None:
            continue
        name = str(row["attribute_name"] or "").strip()
        value = str(row["value_normalized"] or row["value_raw"] or "").strip()
        if name and value and name not in product.attributes:
            product.attributes[name] = value

    return products


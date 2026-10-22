from __future__ import annotations

import sqlite3
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .db_utils import connect_database
from .number_parsing import parse_number


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
    min_shipment: float | None = None
    pack_multiple: float | None = None
    package_min_weight_kg: float | None = None
    package_min_volume_m3: float | None = None
    transport_weight_kg: float | None = None
    transport_volume_m3: float | None = None
    transport_length_mm: float | None = None
    transport_width_mm: float | None = None
    transport_height_mm: float | None = None
    gtin_inner: str = ""
    gtin_outer: str = ""
    attributes: dict[str, str] | None = None


def connect(db_path: Path) -> sqlite3.Connection:
    return connect_database(db_path, readonly=True)


def _scalar(conn: sqlite3.Connection, query: str, params: tuple[Any, ...] = ()) -> Any:
    row = conn.execute(query, params).fetchone()
    return row[0] if row else None


def first_number(*values: Any) -> float | None:
    """First value that is a real number; text such as 'нет' is skipped, not read as 0."""
    for item in values:
        number = parse_number(item)
        if number is not None:
            return number
    return None


def get_stats(db_path: Path) -> DatabaseStats:
    with closing(connect(db_path)) as conn:
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
    with closing(connect(db_path)) as conn:
        rows = conn.execute(
            f"""
            SELECT
                p.article,
                -- An empty or blank value does not hide a filled source lower in the list.
                COALESCE(NULLIF(TRIM(ig.name), ''), NULLIF(TRIM(p.name), ''), '') AS name,
                COALESCE(
                    NULLIF(TRIM(ig.full_name), ''), NULLIF(TRIM(ig.name), ''), NULLIF(TRIM(p.name), ''), ''
                ) AS full_name,
                COALESCE(
                    NULLIF(TRIM(ig.class81_code), ''),
                    (
                        SELECT NULLIF(TRIM(class81_code), '')
                        FROM product_etm_class_suggestions s
                        WHERE s.article = p.article AND NULLIF(TRIM(class81_code), '') IS NOT NULL
                        ORDER BY confidence DESC
                        LIMIT 1
                    ),
                    ''
                ) AS class81_code,
                -- The manufacturer is not a unit: an empty unit becomes "шт" below.
                COALESCE(NULLIF(TRIM(p.unit), ''), '') AS unit,
                d.length_mm,
                d.width_mm,
                d.height_mm,
                -- Main template weight is strictly unit weight. Do not fall back to
                -- min/transport package weights here; those belong to Упак3/Упак5.
                d.weight_kg AS resolved_weight_kg,
                (
                    SELECT NULLIF(psi.gross_weight_unit, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.gross_weight_unit IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS gross_weight_unit,
                (
                    SELECT NULLIF(psi.net_weight_unit, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.net_weight_unit IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS net_weight_unit,
                d.volume_m3 AS resolved_volume_m3,
                (
                    SELECT NULLIF(psi.unit_volume, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.unit_volume IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS unit_volume,
                COALESCE((
                    SELECT tnved_code
                    FROM product_tnved_codes t
                    WHERE t.article = p.article
                    ORDER BY t.import_id DESC, t.id DESC
                    LIMIT 1
                ), '') AS tnved_code,
                COALESCE((
                    SELECT COALESCE(NULLIF(TRIM(okpd2_9), ''), NULLIF(TRIM(okpd2_6), ''), '')
                    FROM product_okpd2_codes o
                    WHERE o.article = p.article
                    ORDER BY o.import_id DESC, o.id DESC
                    LIMIT 1
                ), '') AS okpd2_code,
                (
                    SELECT NULLIF(psi.min_shipment, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.min_shipment IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS min_shipment,
                (
                    SELECT NULLIF(psi.pack_multiple, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.pack_multiple IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS pack_multiple,
                (
                    SELECT NULLIF(psi.gross_weight_min, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.gross_weight_min IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS package_min_weight_kg,
                (
                    SELECT NULLIF(psi.min_volume, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.min_volume IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS package_min_volume_m3,
                (
                    SELECT NULLIF(psi.gross_weight_transport, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.gross_weight_transport IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS gross_weight_transport,
                (
                    SELECT gross_weight_kg
                    FROM material_data_items mdi
                    WHERE mdi.article = p.article
                    ORDER BY mdi.id DESC
                    LIMIT 1
                ) AS material_gross_weight_kg,
                (
                    SELECT NULLIF(psi.transport_volume, '')
                    FROM price_snapshot_items psi
                    WHERE psi.article = p.article AND psi.transport_volume IS NOT NULL
                    ORDER BY psi.snapshot_id DESC, psi.id DESC
                    LIMIT 1
                ) AS transport_volume_m3,
                (
                    SELECT outer_length_mm
                    FROM material_data_items mdi
                    WHERE mdi.article = p.article
                    ORDER BY mdi.id DESC
                    LIMIT 1
                ) AS transport_length_mm,
                (
                    SELECT outer_width_mm
                    FROM material_data_items mdi
                    WHERE mdi.article = p.article
                    ORDER BY mdi.id DESC
                    LIMIT 1
                ) AS transport_width_mm,
                (
                    SELECT outer_height_mm
                    FROM material_data_items mdi
                    WHERE mdi.article = p.article
                    ORDER BY mdi.id DESC
                    LIMIT 1
                ) AS transport_height_mm,
                COALESCE((
                    SELECT gtin_inner
                    FROM material_data_items mdi
                    WHERE mdi.article = p.article
                    ORDER BY mdi.id DESC
                    LIMIT 1
                ), '') AS gtin_inner,
                COALESCE((
                    SELECT gtin_outer
                    FROM material_data_items mdi
                    WHERE mdi.article = p.article
                    ORDER BY mdi.id DESC
                    LIMIT 1
                ), '') AS gtin_outer
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
                weight_kg=first_number(row["resolved_weight_kg"], row["gross_weight_unit"], row["net_weight_unit"]),
                volume_m3=first_number(row["resolved_volume_m3"], row["unit_volume"]),
                tnved_code=str(row["tnved_code"] or ""),
                okpd2_code=str(row["okpd2_code"] or ""),
                min_shipment=parse_number(row["min_shipment"]),
                pack_multiple=parse_number(row["pack_multiple"]),
                package_min_weight_kg=parse_number(row["package_min_weight_kg"]),
                package_min_volume_m3=parse_number(row["package_min_volume_m3"]),
                transport_weight_kg=first_number(row["gross_weight_transport"], row["material_gross_weight_kg"]),
                transport_volume_m3=parse_number(row["transport_volume_m3"]),
                transport_length_mm=row["transport_length_mm"],
                transport_width_mm=row["transport_width_mm"],
                transport_height_mm=row["transport_height_mm"],
                gtin_inner=str(row["gtin_inner"] or ""),
                gtin_outer=str(row["gtin_outer"] or ""),
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

"""File names for ETM uploads.

ETM shows uploaded files by name only, so the name has to say what is inside:
upload_<kind>_<series>_<YYYY-MM-DD>[_vN].csv, where kind is "new" (new cards),
"char" (characteristics of existing products) or "fix" (a correction). Series
come from the products in the file. What needs a look before upload goes to
<name>_check.csv next to it. Latin letters only: ETM's portal is safest
with them. The working copy and the report sit next to the CSV under the same
name.
"""
from __future__ import annotations

import datetime as dt
import re
import sqlite3
from collections import Counter
from collections.abc import Iterable
from contextlib import closing
from pathlib import Path

from .db import connect
from .filler import normalize_article

KINDS = ("new", "char", "fix")
MAX_SERIES = 3
WORDS = {"аксессуары": "acc", "аксессуар": "acc", "принадлежности": "acc"}


def series_names(db_path: Path, articles: Iterable[str]) -> dict[str, str]:
    articles = sorted({article for article in articles if article})
    if not articles:
        return {}
    placeholders = ",".join("?" for _ in articles)
    with closing(connect(db_path)) as conn:
        try:
            rows = conn.execute(
                f"SELECT article, COALESCE(series_name, '') FROM products WHERE article IN ({placeholders})", articles
            ).fetchall()
        except sqlite3.OperationalError:
            return {}
    return {normalize_article(article): series for article, series in rows}


def series_slug(series: str) -> str:
    words = []
    for word in re.split(r"[\s_/]+", series.strip().lower()):
        word = WORDS.get(word, word)
        word = re.sub(r"[^a-z0-9.-]", "", word).strip(".-")
        if word:
            words.append(word)
    return "-".join(words)


def series_label(series: Iterable[str], fallback: str = "") -> str:
    """The products' series, most frequent first; "mix" when there are too many to list."""
    counts = Counter(slug for slug in map(series_slug, series) if slug)
    if not counts:
        return series_slug(fallback) or "chint"
    if len(counts) > MAX_SERIES:
        return "mix"
    return "_".join(slug for slug, _ in counts.most_common())


def upload_csv_path(
    output_dir: Path, kind: str, series: Iterable[str], fallback: str = "", when: dt.date | None = None
) -> Path:
    if kind not in KINDS:
        raise ValueError(f"Неизвестный вид загрузки: {kind}")
    when = when or dt.date.today()
    stem = f"upload_{kind}_{series_label(series, fallback)}_{when.isoformat()}"
    output_dir.mkdir(parents=True, exist_ok=True)
    candidate, version = stem, 1
    while any(output_dir.glob(f"{candidate}.*")) or any(output_dir.glob(f"{candidate}_report.*")):
        version += 1
        candidate = f"{stem}_v{version}"
    return output_dir / f"{candidate}.csv"


def report_path_for(csv_path: Path) -> Path:
    return csv_path.with_name(f"{csv_path.stem}_report.xlsx")


def check_path_for(csv_path: Path) -> Path:
    """Rows or cards held back for review: upload them only after a look."""
    return csv_path.with_name(f"{csv_path.stem}_check{csv_path.suffix}")

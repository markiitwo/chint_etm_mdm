from __future__ import annotations

import argparse
from pathlib import Path

from .analyzer import analyze_template_mapping
from .db import get_stats
from .etim_importer import (
    import_etim_workbook,
    restore_database_backup,
    result_lines as etim_result_lines,
)
from .filler import fill_template
from .price_importer import download_price_file, find_latest_price_url, import_price_workbook, result_lines


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill ETM upload_goods templates from CHINT MDM SQLite.")
    parser.add_argument("--db", required=True, help="Path to chint_mdm.sqlite")
    parser.add_argument("--template", help="Path to upload_goods CSV/XLSX template")
    parser.add_argument("--output-dir", help="Directory for filled file and report")
    parser.add_argument("--rules", help="Optional path to attribute_mappings.json")
    parser.add_argument("--stats", action="store_true", help="Print database status before filling")
    parser.add_argument("--import-price", help="Import a local CHINT price-list XLSX into the database")
    parser.add_argument("--price-url", help="Download and import a CHINT price-list XLSX URL")
    parser.add_argument("--import-etim", help="Import dimensions and attributes from an ETIM XLSX into the database")
    parser.add_argument("--restore-backup", help="Restore selected SQLite .bak file over --db")
    parser.add_argument(
        "--find-latest-price",
        action="store_true",
        help="Find the latest low-voltage CHINT price-list URL on ensmas.ru before importing",
    )
    parser.add_argument("--downloads-dir", help="Directory for downloaded price-list files")
    parser.add_argument(
        "--analyze-mapping",
        action="store_true",
        help="Create XLSX mapping coverage report without filling the template",
    )
    args = parser.parse_args()

    db_path = Path(args.db)
    template_path = Path(args.template) if args.template else None
    output_dir = Path(args.output_dir) if args.output_dir else None
    rules_path = Path(args.rules) if args.rules else None

    if args.restore_backup:
        safety_backup = restore_database_backup(db_path, Path(args.restore_backup))
        print(f"restored_from={args.restore_backup}")
        print(f"previous_db_backup={safety_backup}")
        return

    if args.import_etim:
        result = import_etim_workbook(Path(args.import_etim), db_path, report_dir=output_dir)
        for line in etim_result_lines(result):
            print(line)
        return

    if args.import_price or args.price_url or args.find_latest_price:
        price_path = Path(args.import_price) if args.import_price else None
        price_url = args.price_url or ""
        if args.find_latest_price:
            price_url = find_latest_price_url()
            print(f"price_url={price_url}")
        if price_url:
            downloads_dir = (
                Path(args.downloads_dir)
                if args.downloads_dir
                else (output_dir or Path.cwd()) / "downloads" / "price"
            )
            price_path = download_price_file(price_url, downloads_dir)
        assert price_path is not None
        result = import_price_workbook(price_path, db_path, source_url=price_url)
        for line in result_lines(result):
            print(line)
        return

    if args.stats:
        stats = get_stats(db_path)
        print(f"products={stats.products_count}")
        print(f"dimensions={stats.dimensions_count}")
        print(f"attributes={stats.attributes_count}")
        print(f"latest_price={stats.latest_price_snapshot}")

    if args.analyze_mapping:
        if template_path is None or output_dir is None:
            parser.error("--template and --output-dir are required for --analyze-mapping")
        report_path = analyze_template_mapping(db_path, template_path, output_dir, rules_path)
        print(f"mapping_report={report_path}")
        return

    if template_path is None or output_dir is None:
        parser.error("--template and --output-dir are required for filling")
    result = fill_template(db_path, template_path, output_dir, rules_path)
    print(f"output={result.output_path}")
    print(f"report={result.report_path}")
    print(f"rows={result.total_rows}")
    print(f"found_articles={result.found_articles}")
    print(f"missing_articles={result.missing_articles}")
    print(f"filled_cells={result.filled_cells}")
    print(f"suggested_cells={result.suggested_cells}")
    print(f"missing_cells={result.missing_cells}")


if __name__ == "__main__":
    main()

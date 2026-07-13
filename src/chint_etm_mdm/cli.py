from __future__ import annotations

import argparse
from pathlib import Path

from .analyzer import analyze_template_mapping
from .db import get_stats
from .filler import fill_template


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill ETM upload_goods templates from CHINT MDM SQLite.")
    parser.add_argument("--db", required=True, help="Path to chint_mdm.sqlite")
    parser.add_argument("--template", required=True, help="Path to upload_goods CSV/XLSX template")
    parser.add_argument("--output-dir", required=True, help="Directory for filled file and report")
    parser.add_argument("--rules", help="Optional path to attribute_mappings.json")
    parser.add_argument("--stats", action="store_true", help="Print database status before filling")
    parser.add_argument(
        "--analyze-mapping",
        action="store_true",
        help="Create XLSX mapping coverage report without filling the template",
    )
    args = parser.parse_args()

    db_path = Path(args.db)
    template_path = Path(args.template)
    output_dir = Path(args.output_dir)
    rules_path = Path(args.rules) if args.rules else None

    if args.stats:
        stats = get_stats(db_path)
        print(f"products={stats.products_count}")
        print(f"dimensions={stats.dimensions_count}")
        print(f"attributes={stats.attributes_count}")
        print(f"latest_price={stats.latest_price_snapshot}")

    if args.analyze_mapping:
        report_path = analyze_template_mapping(db_path, template_path, output_dir, rules_path)
        print(f"mapping_report={report_path}")
        return

    result = fill_template(db_path, template_path, output_dir, rules_path)
    print(f"output={result.output_path}")
    print(f"report={result.report_path}")
    print(f"rows={result.total_rows}")
    print(f"found_articles={result.found_articles}")
    print(f"missing_articles={result.missing_articles}")
    print(f"filled_cells={result.filled_cells}")
    print(f"suggested_cells={result.suggested_cells}")


if __name__ == "__main__":
    main()

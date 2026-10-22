from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

from .analyzer import analyze_template_mapping
from .config import incoming_folders, load_config
from .db import get_stats
from .etm_categories import fill_new_products, latest_codes_file, new_articles_since
from .etim_importer import (
    apply_etim_conflict_decisions,
    decision_result_lines,
    import_etim_workbook,
    restore_database_backup,
    result_lines as etim_result_lines,
)
from .export_templates import DEFAULT_PARTNER, available_partners, etm_reference_path
from .filler import fill_template
from .manual_values import (
    import_manual_completions,
    result_lines as manual_result_lines,
)
from .value_matching import import_etm_reference
from .site_info import default_cache_path as site_cache_path
from .certificates import TABLE_NAME as CERTIFICATES_TABLE, table_lines as certificates_lines, update_table as update_certificates_table
from .pm_requests import answer_lines, answers_path, import_answers, latest_assignment_file, load_assignment, request_lines, write_requests
from .raek import fill_raek_zip, remember_name_answers, result_lines as raek_result_lines
from .price_importer import download_price_file, find_latest_price_url, import_price_workbook, result_lines


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill ETM upload_goods templates from CHINT MDM SQLite.")
    parser.add_argument("--db", required=True, help="Path to chint_mdm.sqlite")
    parser.add_argument("--template", help="Path to upload_goods CSV/XLSX template")
    parser.add_argument("--output-dir", help="Directory for filled file and report")
    parser.add_argument(
        "--partner",
        default=DEFAULT_PARTNER,
        choices=available_partners(),
        help="Partner export template to fill (default: etm)",
    )
    parser.add_argument("--rules", help="Optional path to attribute_mappings.json")
    parser.add_argument("--manual-values", help="Optional path to manual_values.json")
    parser.add_argument("--stats", action="store_true", help="Print database status before filling")
    parser.add_argument(
        "--audit-dimensions",
        action="store_true",
        help="List product sizes earlier ETIM imports took from wrong columns (report only, the base is not changed)",
    )
    parser.add_argument("--import-price", help="Import a local CHINT price-list XLSX into the database")
    parser.add_argument("--price-url", help="Download and import a CHINT price-list XLSX URL")
    parser.add_argument("--import-etim", help="Import dimensions and attributes from an ETIM XLSX into the database")
    parser.add_argument("--apply-etim-decisions", help="Apply accepted ETIM conflict decisions from report XLSX")
    parser.add_argument("--import-manual-filled", help="Import manual completions from a filled upload_goods XLSX")
    parser.add_argument(
        "--import-etm-reference",
        help="Import the ETM characteristics reference XLSX (from iPRO OneTeam) next to --rules",
    )
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
    parser.add_argument(
        "--new-since",
        help="Make ETM cards for products introduced after this date (DD.MM.YYYY), one file per class",
    )
    parser.add_argument("--articles", help="Comma-separated articles for new ETM cards instead of --new-since")
    parser.add_argument("--category-templates", help="Folder with ETM category templates (upload_goods_category_template_<class>.xlsx)")
    parser.add_argument("--codes-file", help="Честный знак product list XLSX with ТН ВЭД and ОКПД2; a newer one in Downloads wins")
    parser.add_argument("--raek-zip", help="Fill the RAEK manager's archive (zip) with --output-dir; ЧЗ codes from --codes-file or Downloads")
    parser.add_argument("--raek-site", action="store_true", help="With --raek-zip: first ask ensmas.ru about articles not in the site cache")
    parser.add_argument("--site-cache", help="Site answers cache (default: <work dir>/cache/ensmas_site.json next to --rules)")
    parser.add_argument("--certificates-table", help="Create or update the certificates table (xlsx): dates once per document; ЧЗ codes from --codes-file or Downloads")
    parser.add_argument("--certificate-dates", help="With --certificates-table: take dates from an earlier filled file (a row per product or per document)")
    parser.add_argument("--pm-requests", action="store_true", help="With --raek-zip and --rules: write an Excel request per product manager for what the database lacks")
    parser.add_argument("--pm-answers", nargs="+", help="Import request files sent back by product managers (needs --rules)")
    parser.add_argument("--raek-answers", help="Remember the да/нет answers from the «Названия» sheet of a RAEK report (needs --rules)")
    args = parser.parse_args()

    db_path = Path(args.db)
    template_path = Path(args.template) if args.template else None
    output_dir = Path(args.output_dir) if args.output_dir else None
    rules_path = Path(args.rules) if args.rules else None
    manual_values_file = Path(args.manual_values) if args.manual_values else None

    if args.import_etm_reference:
        if rules_path is None:
            parser.error("--rules is required for --import-etm-reference")
        target = etm_reference_path(rules_path)
        characteristics, values = import_etm_reference(Path(args.import_etm_reference), target)
        print(f"etm_reference={target}")
        print(f"characteristics={characteristics}")
        print(f"allowed_values={values}")
        return

    if args.audit_dimensions:
        from .dimension_audit import audit_dimensions, audit_lines

        report_dir = output_dir or db_path.parent / "reports"
        for line in audit_lines(audit_dimensions(db_path, report_dir)):
            print(line)
        return

    if args.certificates_table:
        table = update_certificates_table(
            db_path,
            latest_codes_file(incoming_folders(load_config().inbox_dir), Path(args.codes_file) if args.codes_file else None),
            Path(args.certificates_table),
            dates_file=Path(args.certificate_dates) if args.certificate_dates else None,
        )
        for line in certificates_lines(table):
            print(line)
        return

    if args.pm_answers:
        if rules_path is None:
            parser.error("--rules is required for --pm-answers")
        for path in args.pm_answers:
            for line in answer_lines(import_answers(Path(path), rules_path.parent.parent)):
                print(line)
        return

    if args.raek_answers:
        if rules_path is None:
            parser.error("--rules is required for --raek-answers")
        kept, rejected = remember_name_answers(Path(args.raek_answers), rules_path)
        print(f"kept={kept}")
        print(f"rejected={rejected}")
        return

    if args.raek_zip:
        if output_dir is None:
            parser.error("--output-dir is required for --raek-zip")
        if args.site_cache:
            site_cache = Path(args.site_cache)
        else:
            site_cache = site_cache_path(rules_path.parent.parent) if rules_path else None
        result = fill_raek_zip(
            db_path,
            Path(args.raek_zip),
            output_dir,
            latest_codes_file(
                incoming_folders(load_config().inbox_dir), Path(args.codes_file) if args.codes_file else None
            ),
            rules_path,
            site_cache,
            args.raek_site,
            progress=print,
            certificates_table=site_cache.parent.parent / CERTIFICATES_TABLE if site_cache else None,
            answers_path=answers_path(rules_path.parent.parent) if rules_path else None,
        )
        for line in raek_result_lines(result):
            print(line)
        if args.pm_requests:
            if rules_path is None:
                parser.error("--rules is required for --pm-requests")
            work = rules_path.parent.parent
            assignment_file = latest_assignment_file(incoming_folders(load_config().inbox_dir) + [work / "input", work])
            if assignment_file is None:
                parser.error("«Закрепление ПМ.xlsx» not found in Downloads or the work folder")
            requests = write_requests(
                result.needs, load_assignment(assignment_file), work, output_dir / "запросы_продактам", f"Заполняем данные для РАЭК ({result.output_path.stem})"
            )
            for line in request_lines(requests) or ["Новых вопросов нет."]:
                print(line)
        return

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

    if args.import_manual_filled:
        if output_dir is None:
            parser.error("--output-dir is required for --import-manual-filled")
        if manual_values_file is None:
            parser.error("--manual-values is required for --import-manual-filled")
        result = import_manual_completions(
            db_path,
            Path(args.import_manual_filled),
            manual_values_file,
            output_dir,
            rules_path,
        )
        for line in manual_result_lines(result):
            print(line)
        return

    if args.apply_etim_decisions:
        result = apply_etim_conflict_decisions(Path(args.apply_etim_decisions), db_path)
        for line in decision_result_lines(result):
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

    if args.new_since or args.articles:
        if output_dir is None or not args.category_templates:
            parser.error("--output-dir and --category-templates are required for new product cards")
        if args.articles:
            articles = [item.strip() for item in args.articles.split(",") if item.strip()]
        else:
            articles = new_articles_since(db_path, dt.datetime.strptime(args.new_since, "%d.%m.%Y").date())
        run = fill_new_products(
            db_path,
            articles,
            Path(args.category_templates),
            output_dir,
            rules_path,
            manual_values_file,
            latest_codes_file(
                incoming_folders(load_config().inbox_dir), Path(args.codes_file) if args.codes_file else None
            ),
        )
        print(f"articles={len(articles)}")
        print(f"summary={run.summary_path}")
        print(f"codes_file={run.codes_file or ''}")
        for code, (result, csv_path) in sorted(run.results.items()):
            print(
                f"class={code} rows={result.found_articles} safe={result.safe_rows} check={result.check_rows} "
                f"missing={result.missing_cells} csv={csv_path} check_csv={result.check_path or ''}"
            )
        for code, members in sorted(run.missing_templates.items()):
            print(f"missing_template={code} articles={len(members)}")
        print(f"unclassified={len(run.unclassified)}")
        print(f"not_found={len(run.not_found)}")
        return

    if args.stats:
        stats = get_stats(db_path)
        print(f"products={stats.products_count}")
        print(f"dimensions={stats.dimensions_count}")
        print(f"attributes={stats.attributes_count}")
        print(f"latest_price={stats.latest_price_snapshot}")
        if template_path is None and output_dir is None and not args.analyze_mapping:
            return  # statistics alone; with --template it is printed before filling

    if args.analyze_mapping:
        if template_path is None or output_dir is None:
            parser.error("--template and --output-dir are required for --analyze-mapping")
        report_path = analyze_template_mapping(db_path, template_path, output_dir, rules_path)
        print(f"mapping_report={report_path}")
        return

    if template_path is None or output_dir is None:
        parser.error("--template and --output-dir are required for filling")
    result = fill_template(
        db_path, template_path, output_dir, rules_path, manual_values_file, partner=args.partner
    )
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

# CHINT ETM MDM

Portable desktop helper for CHINT ETM `upload_goods` templates.

The first MVP is intentionally read-only for the SQLite MDM database:

- opens an existing `chint_mdm.sqlite`;
- shows basic database status;
- fills ETM `upload_goods` CSV/XLSX templates by article;
- saves a filled copy and an XLSX report.

## Run From Source

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python -m chint_etm_mdm
```

For local testing you can point the app at:

```text
/home/openclaw/CHINT-MDM/db/chint_mdm.sqlite
```

## MVP Scope

Filled confidently from the database:

- `Код производителя`
- `Расширенный артикул`
- `81 класс`
- `Название`
- `Полное название`
- `Краткое имя WMS`
- `Страна`
- `Код ТН ВЭД`
- `Код ОКПД2`
- `Название упаковки`
- `Вес, кг`
- `Длина, м`
- `Ширина, м`
- `Высота, м`
- `Объем, м3`
- exact-match `Конфиг:*` fields from ETIM attribute names

The database is not modified by the filler.

For XLSX templates the filler writes and reports only columns whose header cells
are yellow in the template. Non-yellow columns are treated as optional/manual
fields for now. CSV templates do not contain color metadata, so they are filled
by known column names.

The filled XLSX keeps the original dropdown lists/data validations and marks
unfilled required cells in red.

After filling from the GUI, the result window shows a short human-readable
summary: filled cells, red cells, missing articles, the filled file, and the
report. The fill tab also has quick buttons to open the filled workbook, the
latest report, and the output folder.

ETM characteristics are not mapped globally across the whole catalog. A mapping
rule is scoped to an ETM 81 class, for example:

- `class81=50301005`, `Конфиг:Напряжение, В` -> `Напряжение лампы, В`;
- `class81=50301005`, `Конфиг:Цвет свечения` -> `Цвет`.

Only `approved_class_rule` mappings are written into templates. Candidate
matches are kept in the mapping analysis report until a human confirms the
meaning for that class.

On first workspace setup the app creates an editable rules file:

```text
<work-dir>/rules/attribute_mappings.json
```

If this file exists, it controls class-scoped mappings. This lets you add or
remove approved mappings without rebuilding the EXE.

## CLI Check

The same filler can be run without the GUI:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --template /path/to/upload_goods.xlsx \
  --output-dir /path/to/output \
  --stats
```

The command creates:

- a filled copy of the template;
- an Excel report with summary, detailed actions, and a `К продактам`
  sheet containing unfilled fields by article.

To inspect characteristic coverage without changing a template, run:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --template /path/to/upload_goods_category_template.xlsx \
  --output-dir ./output \
  --rules /path/to/work-dir/rules/attribute_mappings.json \
  --analyze-mapping
```

This creates `*_mapping_review_*.xlsx` with class-scoped coverage and candidate
attributes. Use it to decide which mappings can be safely approved for a class.

In the GUI, open the `Правила маппинга` tab, choose the generated
`mapping_review.xlsx`, and load candidates. The upper table shows coverage for
each yellow field: what will be filled, what needs a source choice, and what
should be sent to product managers. The lower table shows candidate source
attributes, sample source values, and how many products this source can cover.
Use `Принять` only when the source is correct for the product category. Use
`Отклонить` for noisy candidates so they stop appearing in future mapping
reports. Batch mode is also available through the checkbox column and the
`Сохранить выбранные правила` / `Отклонить выбранные` buttons.

The mapping review workbook also includes:

- `Покрытие` — field-level status for all yellow columns;
- `К продактам` — unfilled fields by article, ready to send for enrichment;
- `Выбор источника` — fields that have source candidates but require a human
  decision first;
- `Правила` — candidates that can be accepted or rejected in the GUI.

The app writes approved and rejected class-scoped decisions to:

```text
<work-dir>/rules/attribute_mappings.json
```

The next fill run uses approved class rules immediately. Rejected candidates are
ignored by the next mapping analysis. Rebuilding the EXE is not required.

Report statuses:

- `filled` means the value was written into the template;
- `candidate` means a close ETIM attribute was found for review, but it is not written until approved;
- `not_found` means the article was not found in the database.

## v1 Workflow For Colleagues

1. Choose the SQLite database and an `upload_goods` XLSX template.
2. Click `Проанализировать маппинг` when the template has unknown
   `Конфиг:` fields. Accept only sources that clearly match the field meaning;
   reject noisy sources so they stop appearing later.
3. Click `Заполнить из базы`.
4. Open the generated `_filled.xlsx`. Values filled by the app stay in the
   template, original dropdown lists remain available, and required fields that
   still need manual/product-manager input are red.
5. Send the report sheet `К продактам` or the red cells in `_filled.xlsx` for
   enrichment.

The v1 app does not write manual corrections back into the SQLite database.
That is planned as a separate v2 flow so the source database remains easy to
restore and audit.

## v2 Price Update

The GUI tab `Обновление базы` can update the selected SQLite database from a
fresh CHINT price-list workbook:

1. Choose a local `Price-list-CHINT_*.xlsx` file, or paste a direct URL and let
   the app download it into `<work-dir>/downloads/price`.
2. Click `Импортировать выбранный XLSX` or `Скачать и импортировать`.
3. The app creates a timestamped `.bak_*` copy of the SQLite database before
   writing.
4. The import updates `products`, writes a new `price_snapshots` row, appends
   `price_snapshot_items`, and imports the workbook sheets for new products,
   outgoing assortment, and price-change history.
5. The result window shows imported rows, new articles compared with the
   previous snapshot, changed price rows, total products, and backup path.

CLI import is also available:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --import-price /path/to/Price-list-CHINT_01-01-2026.xlsx
```

Or download first:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --price-url https://example.com/Price-list-CHINT.xlsx \
  --downloads-dir /path/to/work-dir/downloads/price
```

## Build EXE

Planned packaging command:

```bash
python -m PyInstaller --noconfirm --windowed --onedir --name "CHINT ETM MDM" --paths src run_app.py
```

Packaging will be refined after the MVP is tested on Windows.

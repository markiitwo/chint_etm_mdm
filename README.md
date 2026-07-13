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
- an Excel report with summary and detailed action sheets.

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

Report statuses:

- `filled` means the value was written into the template;
- `filled_suggested` means a close ETIM attribute was found, written into the template, and marked for review;
- `not_found` means the article was not found in the database.

## Build EXE

Planned packaging command:

```bash
python -m PyInstaller --noconfirm --windowed --onedir --name "CHINT ETM MDM" --paths src run_app.py
```

Packaging will be refined after the MVP is tested on Windows.

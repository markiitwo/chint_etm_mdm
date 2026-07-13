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

Some ETM characteristic names are mapped to source attribute aliases. For
example, `Конфиг:Напряжение, В` can be filled from `Напряжение лампы, В`.

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

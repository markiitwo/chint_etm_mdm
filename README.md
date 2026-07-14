# CHINT ETM MDM

Приложение для подготовки ETM `upload_goods`: оно заполняет шаблоны по базе
CHINT, показывает что не удалось заполнить, помогает принять правила маппинга и
обновлять локальную SQLite-базу из прайса и ETIM-файлов.

## Что умеет

- заполнять `upload_goods.xlsx` по артикулам;
- сохранять выпадающие списки из исходного шаблона;
- подсвечивать незаполненные обязательные ячейки красным;
- делать понятный отчет для продакт-менеджеров;
- анализировать `Конфиг:*` поля и сохранять подтвержденные правила маппинга;
- скачивать свежий прайс-лист CHINT с `ensmas.ru` и импортировать его в базу;
- импортировать ETIM-файл, добирая характеристики и габариты;
- разбирать конфликты ETIM через Excel-отчет с колонкой `Решение`;
- откатывать базу из `.bak_*`;
- запоминать ручные дозаполнения из `_filled.xlsx` и использовать их при
  следующем заполнении.

## Обычный сценарий

1. На вкладке `База` выберите рабочую папку и SQLite-базу.
2. На вкладке `Заполнение upload_goods` выберите шаблон и папку результата.
3. Нажмите `Проанализировать маппинг`, если в шаблоне есть незнакомые
   `Конфиг:*` поля.
4. На вкладке `Правила маппинга` принимайте только те источники, которые точно
   подходят по смыслу. Мусорные варианты отклоняйте.
5. Нажмите `Заполнить из базы`.
6. Откройте `_filled.xlsx`: незаполненные обязательные поля будут красными.
7. Лист `К продактам` из отчета можно отправлять на дозаполнение.
8. Если файл дозаполнили руками, импортируйте его в блоке
   `Ручные дозаполнения`. Следующее заполнение уже подхватит эти значения.

## Обновление базы

Все операции, которые меняют SQLite-базу, сначала создают бэкап рядом с базой.
Если что-то пошло не так, на вкладке `База` можно выбрать `.bak_*` и нажать
`Откатить базу из бэкапа`.

### Прайс-лист

На вкладке `Обновление базы`:

1. нажмите `Найти свежий прайс`, чтобы программа нашла актуальный файл на
   `ensmas.ru`;
2. нажмите `Скачать и импортировать`;
3. после импорта проверьте статистику в журнале.

Можно также выбрать уже скачанный `Price-list-CHINT_*.xlsx` и нажать
`Импортировать выбранный XLSX`.

### ETIM-файл

1. В блоке `ETIM-файл` выберите `.xlsx/.xlsm`.
2. Нажмите `Импортировать ETIM`.
3. Программа найдет артикулы из базы во всех листах ETIM-файла, возьмет
   характеристики и попробует дозаполнить пустые габариты.
4. Уже заполненные габариты не перетираются автоматически. Если ETIM дает
   другое значение, строка попадает в отчет `Конфликты`.
5. Откройте отчет кнопкой `Открыть отчет ETIM`.
6. На листе `Конфликты` выберите в колонке `Решение`:
   - `Оставить базу` — ничего не менять;
   - `Принять ETIM` — записать значение из ETIM в базу.
7. Сохраните отчет и нажмите `Применить решения ETIM`.

## Ручные дозаполнения

Это отдельный слой поверх базы. Он нужен, чтобы не терять ручную работу и не
заставлять людей каждый раз заполнять одно и то же.

1. Заполните шаблон через приложение.
2. Передайте `_filled.xlsx` на ручное дозаполнение.
3. Выберите дозаполненный файл в блоке `Ручные дозаполнения`.
4. Нажмите `Импортировать ручные значения`.
5. Программа сохранит отличия в:

```text
<рабочая папка>/rules/manual_values.json
```

При следующем `Заполнить из базы` эти значения будут применены поверх данных из
SQLite.

## Где лежат важные файлы

```text
<рабочая папка>/database/chint_mdm.sqlite       рабочая копия базы
<рабочая папка>/output/                         заполненные файлы и отчеты
<рабочая папка>/reports/                        отчеты импортов
<рабочая папка>/downloads/price/                скачанные прайсы
<рабочая папка>/rules/attribute_mappings.json   правила маппинга
<рабочая папка>/rules/manual_values.json        ручные дозаполнения
```

## Запуск из исходников

```bash
python3 -m venv .venv
. .venv/bin/activate
pip install -r requirements.txt
python -m chint_etm_mdm
```

## CLI

Заполнить шаблон:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --template /path/to/upload_goods.xlsx \
  --output-dir /path/to/work/output \
  --rules /path/to/work/rules/attribute_mappings.json \
  --manual-values /path/to/work/rules/manual_values.json
```

Проанализировать маппинг:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --template /path/to/upload_goods.xlsx \
  --output-dir /path/to/work/output \
  --analyze-mapping
```

Импортировать свежий прайс:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --find-latest-price \
  --downloads-dir /path/to/work/downloads/price
```

Импортировать ETIM:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --import-etim /path/to/etim.xlsx \
  --output-dir /path/to/work/reports
```

Применить решения ETIM:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --apply-etim-decisions /path/to/etim_import_report.xlsx
```

Импортировать ручные дозаполнения:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --import-manual-filled /path/to/upload_goods_filled.xlsx \
  --output-dir /path/to/work/output \
  --manual-values /path/to/work/rules/manual_values.json
```

Откатить базу:

```bash
PYTHONPATH=src python -m chint_etm_mdm.cli \
  --db /path/to/chint_mdm.sqlite \
  --restore-backup /path/to/chint_mdm.sqlite.bak_2026-07-14_12-00-00
```

## Сборка EXE

```bash
python -m PyInstaller --noconfirm --windowed --onedir --name "CHINT ETM MDM" --paths src run_app.py
```

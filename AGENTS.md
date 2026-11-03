# Инструкции для агентов

CHINT ETM MDM — десктоп-программа (PySide6 + openpyxl + SQLite): готовит карточки товаров CHINT для ЭТМ и РАЭК.
Код — `src/chint_etm_mdm/`, тесты — `tests/` (unittest, pytest не используется).

## Проверки

```bash
PYTHONPATH=src QT_QPA_PLATFORM=offscreen .venv/bin/python -m unittest discover -s tests
PYTHONPATH=src QT_QPA_PLATFORM=offscreen .venv/bin/python -m chint_etm_mdm.app --self-test
ruff check --select F src tests   # то же в CI
```

## Правила

- Рабочую базу (`~/CHINT ETM MDM Work/database/chint_mdm.sqlite`) и папку `~/CHINT ETM MDM Work` не открывать на запись.
  Замеры — на копии в `/tmp` или с `?mode=ro`.
- Тексты интерфейса и сообщений (русские строки) не менять без запроса — они согласованы с пользователями.
- Версия — в `src/chint_etm_mdm/__init__.py`, правило номеров — в шапке `CHANGELOG.md`.
- Минимальный дифф, без новых зависимостей и кода «на будущее». Новая логика — небольшой тест.
- Перенося код, сохраняй его обработку ошибок и проверки.
- План рефакторинга и статус — `docs/REFACTOR_PLAN.md`. Старые планы в `docs/` не читать.

## Модули

- `app.py` — GUI (вкладки, воркеры, диалоги); `__main__.py` — запуск.
- `config.py` — настройки и пути по умолчанию; `inputs.py` — копии входных файлов в рабочей папке.
- `db.py` — чтение товаров (`fetch_products`, `get_stats`); `db_utils.py` — подключение, бэкапы, копирование базы.
- `price_importer.py` — прайс; `etim_importer.py` — ETIM-характеристики; `dimension_audit.py` — проверка размеров.
- `manual_values.py` — ручные дополнения; `mapping_rules.py` — правила сопоставления полей.
- `analyzer.py` — отчёт маппинга шаблона; `filler.py` — заполнение шаблона партнёра.
- `export_templates.py` — шаблоны партнёров; `value_matching.py` — значения под справочник партнёра.
- `etm_categories.py` — новые карточки ЭТМ по шаблонам классов; `etm_characteristics.py` — отчёт недостающих характеристик.
- `etm_materials.py` — фото, документы, сертификаты, описания для ЭТМ; `upload_names.py` — имена файлов выгрузки.
- `raek.py` — архив РАЭК; `certificates.py` — даты сертификатов; `site_info.py` — данные с ensmas.ru.
- `pm_requests.py` — запросы продактам и их ответы; `data_checks.py` — задержка сомнительных значений.
- `excel_files.py`, `atomic_files.py`, `number_parsing.py` — общие помощники.
- `updater.py`, `update_ui.py` — обновления сборок; `notify.py` — почта и Telegram; `net.py` — открытие ссылок.

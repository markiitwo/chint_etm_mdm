from __future__ import annotations

import shutil
import sys
import traceback
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .config import AppConfig, ensure_work_dirs, load_config, save_config
from .db import get_stats
from .filler import FillResult, fill_template


class FillWorker(QThread):
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, db_path: Path, template_path: Path, output_dir: Path) -> None:
        super().__init__()
        self.db_path = db_path
        self.template_path = template_path
        self.output_dir = output_dir

    def run(self) -> None:
        try:
            self.done.emit(fill_template(self.db_path, self.template_path, self.output_dir))
        except Exception:
            self.failed.emit(traceback.format_exc())


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("CHINT ETM MDM")
        self.resize(980, 680)
        self.config = load_config()
        self.worker: FillWorker | None = None

        self.work_dir_edit = QLineEdit(self.config.work_dir)
        self.db_path_edit = QLineEdit(self.config.db_path)
        self.template_path_edit = QLineEdit()
        self.output_dir_edit = QLineEdit(self.default_output_dir())
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.fill_log = QTextEdit()
        self.fill_log.setReadOnly(True)

        self.setCentralWidget(self.build_ui())
        self.refresh_database_status()

    def build_ui(self) -> QWidget:
        tabs = QTabWidget()
        tabs.addTab(self.build_database_tab(), "База")
        tabs.addTab(self.build_fill_tab(), "Заполнение upload_goods")

        root = QWidget()
        layout = QVBoxLayout(root)
        layout.addWidget(tabs)
        return root

    def build_database_tab(self) -> QWidget:
        root = QWidget()
        layout = QVBoxLayout(root)

        setup_box = QGroupBox("Рабочая папка и база")
        form = QGridLayout(setup_box)
        form.addWidget(QLabel("Рабочая папка"), 0, 0)
        form.addWidget(self.work_dir_edit, 0, 1)
        choose_work = QPushButton("Выбрать...")
        choose_work.clicked.connect(self.choose_work_dir)
        form.addWidget(choose_work, 0, 2)

        form.addWidget(QLabel("База SQLite"), 1, 0)
        form.addWidget(self.db_path_edit, 1, 1)
        choose_db = QPushButton("Выбрать...")
        choose_db.clicked.connect(self.choose_db)
        form.addWidget(choose_db, 1, 2)

        buttons = QHBoxLayout()
        save_button = QPushButton("Сохранить настройки")
        save_button.clicked.connect(self.save_current_config)
        backup_button = QPushButton("Копировать базу в рабочую папку")
        backup_button.clicked.connect(self.copy_db_to_workspace)
        refresh_button = QPushButton("Обновить статус")
        refresh_button.clicked.connect(self.refresh_database_status)
        buttons.addWidget(save_button)
        buttons.addWidget(backup_button)
        buttons.addWidget(refresh_button)

        layout.addWidget(setup_box)
        layout.addLayout(buttons)
        layout.addWidget(QLabel("Статус базы"))
        layout.addWidget(self.status_text, stretch=1)
        return root

    def build_fill_tab(self) -> QWidget:
        root = QWidget()
        layout = QVBoxLayout(root)

        box = QGroupBox("Шаблон")
        form = QFormLayout(box)

        template_row = QHBoxLayout()
        template_row.addWidget(self.template_path_edit)
        choose_template = QPushButton("Выбрать...")
        choose_template.clicked.connect(self.choose_template)
        template_row.addWidget(choose_template)
        form.addRow("Файл upload_goods", template_row)

        output_row = QHBoxLayout()
        output_row.addWidget(self.output_dir_edit)
        choose_output = QPushButton("Выбрать...")
        choose_output.clicked.connect(self.choose_output_dir)
        output_row.addWidget(choose_output)
        form.addRow("Папка результата", output_row)

        run_button = QPushButton("Заполнить из базы")
        run_button.clicked.connect(self.run_fill)

        layout.addWidget(box)
        layout.addWidget(run_button, alignment=Qt.AlignLeft)
        layout.addWidget(QLabel("Журнал"))
        layout.addWidget(self.fill_log, stretch=1)
        return root

    def default_output_dir(self) -> str:
        if self.config.work_dir:
            return str(Path(self.config.work_dir) / "output")
        return ""

    def choose_work_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите рабочую папку")
        if not path:
            return
        work_dir = Path(path)
        ensure_work_dirs(work_dir)
        self.work_dir_edit.setText(str(work_dir))
        self.output_dir_edit.setText(str(work_dir / "output"))
        self.save_current_config()

    def choose_db(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите chint_mdm.sqlite",
            "",
            "SQLite database (*.sqlite *.sqlite3 *.db);;All files (*)",
        )
        if path:
            self.db_path_edit.setText(path)
            self.save_current_config()
            self.refresh_database_status()

    def choose_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите upload_goods шаблон",
            "",
            "ETM templates (*.xlsx *.xlsm *.csv);;All files (*)",
        )
        if path:
            self.template_path_edit.setText(path)

    def choose_output_dir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Выберите папку результата")
        if path:
            self.output_dir_edit.setText(path)

    def save_current_config(self) -> None:
        self.config = AppConfig(
            work_dir=self.work_dir_edit.text().strip(),
            db_path=self.db_path_edit.text().strip(),
        )
        if self.config.work_path:
            ensure_work_dirs(self.config.work_path)
        save_config(self.config)

    def copy_db_to_workspace(self) -> None:
        source = Path(self.db_path_edit.text().strip())
        work_dir_text = self.work_dir_edit.text().strip()
        if not source.exists():
            QMessageBox.warning(self, "База не найдена", "Сначала выберите существующую базу.")
            return
        if not work_dir_text:
            QMessageBox.warning(self, "Нет рабочей папки", "Сначала выберите рабочую папку.")
            return
        target_dir = Path(work_dir_text) / "database"
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / "chint_mdm.sqlite"
        shutil.copy2(source, target)
        self.db_path_edit.setText(str(target))
        self.save_current_config()
        self.refresh_database_status()
        QMessageBox.information(self, "Готово", f"База скопирована:\n{target}")

    def refresh_database_status(self) -> None:
        db_text = self.db_path_edit.text().strip()
        if not db_text:
            self.status_text.setPlainText("База не выбрана.")
            return
        db_path = Path(db_text)
        if not db_path.exists():
            self.status_text.setPlainText(f"База не найдена:\n{db_path}")
            return
        try:
            stats = get_stats(db_path)
        except Exception as exc:
            self.status_text.setPlainText(f"Не удалось прочитать базу:\n{exc}")
            return

        self.status_text.setPlainText(
            "\n".join(
                [
                    f"Файл: {db_path}",
                    f"Товаров: {stats.products_count}",
                    f"Габариты/вес: {stats.dimensions_count}",
                    f"ETIM-значений: {stats.attributes_count}",
                    f"Последний прайс: {stats.latest_price_snapshot}",
                    f"Файл прайса: {stats.latest_price_file}",
                ]
            )
        )

    def run_fill(self) -> None:
        self.save_current_config()
        db_path = Path(self.db_path_edit.text().strip())
        template_path = Path(self.template_path_edit.text().strip())
        output_dir = Path(self.output_dir_edit.text().strip())

        if not db_path.exists():
            QMessageBox.warning(self, "База не найдена", "Выберите существующую SQLite базу.")
            return
        if not template_path.exists():
            QMessageBox.warning(self, "Шаблон не найден", "Выберите файл шаблона.")
            return
        if not output_dir:
            QMessageBox.warning(self, "Нет папки результата", "Выберите папку результата.")
            return

        self.fill_log.append("Запуск заполнения...")
        self.worker = FillWorker(db_path, template_path, output_dir)
        self.worker.done.connect(self.on_fill_done)
        self.worker.failed.connect(self.on_fill_failed)
        self.worker.start()

    def on_fill_done(self, result: FillResult) -> None:
        self.fill_log.append("Готово.")
        self.fill_log.append(f"Строк в шаблоне: {result.total_rows}")
        self.fill_log.append(f"Найдено артикулов: {result.found_articles}")
        self.fill_log.append(f"Не найдено артикулов: {result.missing_articles}")
        self.fill_log.append(f"Заполнено ячеек: {result.filled_cells}")
        self.fill_log.append(f"Предложений в отчете: {result.suggested_cells}")
        self.fill_log.append(f"Файл: {result.output_path}")
        self.fill_log.append(f"Отчет: {result.report_path}")
        QMessageBox.information(
            self,
            "Заполнение завершено",
            f"Файл:\n{result.output_path}\n\nОтчет:\n{result.report_path}",
        )

    def on_fill_failed(self, details: str) -> None:
        self.fill_log.append("Ошибка заполнения.")
        self.fill_log.append(details)
        QMessageBox.critical(self, "Ошибка", details)


def main() -> None:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

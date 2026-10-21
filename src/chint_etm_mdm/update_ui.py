"""Install and update dialogs for the ready-made builds (see updater.py)."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QProgressDialog

from . import __version__, updater


class UpdateCheckWorker(QThread):
    found = Signal(object)
    failed = Signal(str)

    def run(self) -> None:
        try:
            self.found.emit(updater.latest_release())
        except Exception as exc:
            self.failed.emit(str(exc))


class UpdateInstallWorker(QThread):
    progress = Signal(int, int)
    done = Signal(object)
    failed = Signal(str)

    def __init__(self, release: updater.Release) -> None:
        super().__init__()
        self.release = release

    def run(self) -> None:
        try:
            self.done.emit(updater.install_release(self.release, progress=self.progress.emit))
        except Exception as exc:
            self.failed.emit(str(exc))


def offer_install() -> bool:
    """Ask to install a portable build that was started from an unpacked archive.

    Returns True when the installed copy was started and this one should quit.
    """
    if not updater.is_frozen() or updater.running_installed():
        return False
    answer = QMessageBox.question(
        None,
        "Установка программы",
        "Установить программу на этот компьютер?\n\n"
        "Появится значок на рабочем столе, а новые версии будут приходить сами. "
        "Ваши базы, шаблоны и настройки не затрагиваются.",
    )
    if answer != QMessageBox.StandardButton.Yes:
        return False
    try:
        exe = updater.install_running_copy()
    except Exception as exc:
        QMessageBox.critical(None, "Установка не удалась", str(exc))
        return False
    updater.launch(exe)
    QMessageBox.information(
        None, "Программа установлена", "Готово. Дальше запускайте программу значком на рабочем столе."
    )
    return True


class Updates:
    """The 'Программа' menu and the quiet update check at start."""

    def __init__(self, window: QMainWindow) -> None:
        self.window = window
        self.worker: QThread | None = None
        self.progress: QProgressDialog | None = None
        menu = window.menuBar().addMenu("Программа")
        menu.addAction("Проверить обновления", lambda: self.check(quiet=False))
        self.rollback_action = menu.addAction("Вернуть предыдущую версию", self.rollback)
        menu.addAction("О программе", self.about)
        menu.aboutToShow.connect(self.refresh_menu)
        if updater.running_installed():
            updater.cleanup_old_versions()
            self.check(quiet=True)

    def refresh_menu(self) -> None:
        self.rollback_action.setEnabled(updater.running_installed() and bool(updater.previous_version()))

    def about(self) -> None:
        where = "установлена" if updater.running_installed() else "запущена без установки"
        QMessageBox.information(self.window, "О программе", f"CHINT ETM MDM, версия {__version__}\nПрограмма {where}.")

    def check(self, quiet: bool) -> None:
        if self.worker is not None:
            return
        if not updater.running_installed():
            if not quiet:
                QMessageBox.information(
                    self.window,
                    "Обновления",
                    "Обновления работают у установленной программы. "
                    "Скачайте готовую версию со страницы выпусков и запустите её: она предложит установиться.",
                )
            return
        worker = UpdateCheckWorker()
        worker.found.connect(lambda release: self.on_checked(release, quiet))
        worker.failed.connect(lambda error: self.on_check_failed(error, quiet))
        self.start(worker)

    def start(self, worker: QThread) -> None:
        self.worker = worker
        worker.finished.connect(self.on_finished)
        worker.start()

    def on_finished(self) -> None:
        self.worker = None

    def on_check_failed(self, error: str, quiet: bool) -> None:
        if not quiet:
            QMessageBox.warning(self.window, "Обновления", f"Не удалось проверить обновления:\n{error}")

    def on_checked(self, release: updater.Release | None, quiet: bool) -> None:
        if release is None or not updater.is_newer(release.version):
            if not quiet:
                QMessageBox.information(self.window, "Обновления", f"У вас последняя версия ({__version__}).")
            return
        answer = QMessageBox.question(
            self.window,
            "Есть новая версия",
            f"Вышла версия {release.version} (у вас {__version__}). Обновить сейчас?\n\n"
            "Программа перезапустится. Если что-то пойдёт не так, в меню «Программа» "
            "можно вернуть предыдущую версию.",
        )
        if answer == QMessageBox.StandardButton.Yes:
            self.install(release)

    def install(self, release: updater.Release) -> None:
        self.progress = QProgressDialog("Скачиваю новую версию...", "", 0, 0, self.window)
        self.progress.setCancelButton(None)
        self.progress.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress.setMinimumDuration(0)
        self.progress.show()
        worker = UpdateInstallWorker(release)
        worker.progress.connect(self.on_progress)
        worker.done.connect(self.on_installed)
        worker.failed.connect(self.on_install_failed)
        self.start(worker)

    def on_progress(self, done: int, total: int) -> None:
        if self.progress and total:
            self.progress.setMaximum(total // 1024)
            self.progress.setValue(done // 1024)

    def on_install_failed(self, error: str) -> None:
        if self.progress:
            self.progress.close()
        QMessageBox.warning(self.window, "Обновление не удалось", f"Работаем на прежней версии.\n\n{error}")

    def on_installed(self, exe: Path) -> None:
        if self.progress:
            self.progress.close()
        self.restart_into(exe)

    def rollback(self) -> None:
        previous = updater.previous_version()
        if not previous:
            return
        answer = QMessageBox.question(
            self.window, "Вернуть предыдущую версию", f"Вернуться к версии {previous}? Программа перезапустится."
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            exe = updater.rollback()
        except Exception as exc:
            QMessageBox.warning(self.window, "Не получилось", str(exc))
            return
        self.restart_into(exe)

    def restart_into(self, exe: Path) -> None:
        busy = getattr(self.window, "worker", None) is not None
        if busy:
            QMessageBox.information(
                self.window,
                "Перезапуск позже",
                "Сейчас идёт операция. Новая версия запустится, когда вы закроете программу и откроете её снова.",
            )
            return
        updater.launch(exe)
        QApplication.quit()

"""Files the program reads live in the work folder: the place a file was saved to (Downloads) is not safe to work from.

A chosen or found file is copied once into a folder of its kind; the program then uses the copy and writes its path
into the settings. The original is never touched.
"""
from __future__ import annotations

import filecmp
import os
import shutil
from itertools import count
from pathlib import Path

# kind of a file -> its folder inside the work folder («прайс» is where the program saves downloaded price lists)
FOLDERS = {
    "ЧЗ": "input/ЧЗ",
    "закрепление ПМ": "input/закрепление_ПМ",
    "ETIM": "input/ETIM",
    "прайс": "downloads/price",
    "шаблон ЭТМ": "input/шаблон_ЭТМ",
    "РАЭК": "input/РАЭК",
    "ответы продактов": "input/ответы_продактов",
    "справочник ЭТМ": "input/справочник_ЭТМ",
}


def kind_folder(work_dir: Path | str, kind: str) -> Path:
    return Path(work_dir) / FOLDERS[kind]


def input_folders(work_dir: Path | str) -> list[Path]:
    """Where the program looks for files it already brought in: `input` and the folders of the kinds inside it."""
    root = Path(work_dir) / "input"
    return [root] + (sorted(path for path in root.iterdir() if path.is_dir()) if root.is_dir() else [])


def bring_in(path: Path | str | None, work_dir: Path | str, kind: str) -> Path | None:
    """The file to work with: `path` itself when it is not a file or already in the work folder, else its copy there.

    The same content is copied once. A file with the same name but other content gets "_1", "_2" after the name
    (the price import already ignores that suffix), nothing is overwritten. The copy keeps the file's date, so
    "the newest file" still means the same file.
    shortcut: old copies are never deleted, add a cleanup when the folders grow.
    """
    if not path:
        return None
    path = Path(path)
    if not work_dir or not path.is_file() or Path(work_dir).resolve() in path.resolve().parents:
        return path
    folder = kind_folder(work_dir, kind)
    folder.mkdir(parents=True, exist_ok=True)
    for n in count():
        name = path.name if n == 0 else f"{path.stem}_{n}{path.suffix}"
        target = folder / name
        if not target.exists():
            temp = folder / f".{name}.tmp"
            try:
                shutil.copy2(path, temp)
                os.replace(temp, target)
            finally:
                temp.unlink(missing_ok=True)
            return target
        if filecmp.cmp(path, target, shallow=False):
            return target

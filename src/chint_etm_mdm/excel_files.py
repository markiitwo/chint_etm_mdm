"""Excel files that come from colleagues in whatever form Excel saved them.

openpyxl reads only the modern zip-based workbook. A file named .xlsx can still
be an old binary Excel book (.xls renamed or saved in compatibility mode), a
workbook protected with a password or a confidentiality label (Excel stores
those encrypted, not as a zip), an HTML table, or a half-downloaded file. The
program recognises each case: old books are converted to .xlsx by LibreOffice
when it is installed, the rest get a plain explanation of what to do.
"""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ZIP_MAGIC = b"PK\x03\x04"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ENCRYPTED_MARKER = "EncryptionInfo".encode("utf-16-le")
CONVERT_TIMEOUT = 600
CONVERTED_DIR = "chint_etm_mdm_xlsx"


def find_soffice() -> str:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    if sys.platform == "win32":
        for root in (os.environ.get("ProgramFiles", ""), os.environ.get("ProgramFiles(x86)", "")):
            candidate = Path(root) / "LibreOffice" / "program" / "soffice.exe"
            if root and candidate.exists():
                return str(candidate)
    return ""


def file_kind(path: Path) -> str:
    """zip | encrypted | ole | html | empty | unknown."""
    with path.open("rb") as handle:
        head = handle.read(8)
        if not head:
            return "empty"
        if head.startswith(ZIP_MAGIC):
            return "zip"
        if head == OLE_MAGIC:
            handle.seek(0)
            return "encrypted" if ENCRYPTED_MARKER in handle.read() else "ole"
    text = head.lstrip(b"\xef\xbb\xbf \t\r\n").lower()
    return "html" if text.startswith(b"<") else "unknown"


def system_env() -> dict[str, str]:
    """Environment for an outside program: the built app's own library path must not leak into LibreOffice."""
    env = dict(os.environ)
    if getattr(sys, "frozen", False):
        original = env.pop("LD_LIBRARY_PATH_ORIG", None)
        if original is None:
            env.pop("LD_LIBRARY_PATH", None)
        else:
            env["LD_LIBRARY_PATH"] = original
    return env


def convert_to_xlsx(path: Path, out_dir: Path, soffice: str) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / f"{path.stem}.xlsx"
    target.unlink(missing_ok=True)
    profile = Path(tempfile.mkdtemp(prefix="chint-lo-"))  # a separate profile works even when LibreOffice is open
    try:
        subprocess.run(
            [soffice, f"-env:UserInstallation={profile.as_uri()}", "--headless", "--convert-to", "xlsx", "--outdir", str(out_dir), str(path)],
            check=False,
            capture_output=True,
            timeout=CONVERT_TIMEOUT,
            env=system_env(),
        )
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if not target.exists() or file_kind(target) != "zip":
        raise ValueError(
            f"Файл «{path.name}» в старом формате Excel, и LibreOffice не смог его пересохранить. "
            "Откройте файл и сохраните его как «Книга Excel (.xlsx)»."
        )
    return target


def readable_xlsx(path: Path, what: str = "Файл", out_dir: Path | None = None, soffice: str | None = None) -> Path:
    """A path openpyxl can open: the file itself, or its .xlsx copy made by LibreOffice."""
    kind = file_kind(path)
    if kind == "zip":
        return path
    if kind == "empty":
        raise ValueError(f"{what} «{path.name}» пустой. Возможно, он ещё не докачался или не скопировался до конца.")
    if kind == "encrypted":
        raise ValueError(
            f"{what} «{path.name}» защищён паролем или меткой конфиденциальности, программа не может его прочитать. "
            "Откройте его в Excel, снимите защиту (или «Сохранить как» без пароля) и сохраните обычной книгой .xlsx."
        )
    if kind in {"ole", "html"}:
        soffice = find_soffice() if soffice is None else soffice
        if soffice:
            return convert_to_xlsx(path, out_dir or Path(tempfile.gettempdir()) / CONVERTED_DIR, soffice)
        raise ValueError(
            f"{what} «{path.name}» сохранён в старом формате Excel. "
            "Откройте его и сохраните как «Книга Excel (.xlsx)», либо установите LibreOffice: тогда программа пересохранит его сама."
        )
    raise ValueError(f"{what} «{path.name}» не похож на книгу Excel. Проверьте, что выбран нужный файл.")

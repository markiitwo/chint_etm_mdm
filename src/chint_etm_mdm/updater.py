"""Install, update and roll back the ready-made program builds published on GitHub.

Layout of an installed program (one folder per version, two kept):

    <install root>/versions/<version>/CHINT ETM MDM[.exe]
    <install root>/state.json        {"current": "0.2.0", "previous": "0.1.2"}

Desktop and menu icons point straight at the current version's executable,
so switching versions means rewriting the icons. No Qt here: the module also
runs from the command line to move a source install onto the release builds.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

from . import __version__
from .net import open_url

APP_NAME = "CHINT ETM MDM"
REPO = "markiitwo/chint_etm_mdm"
LATEST_RELEASE_URL = f"https://api.github.com/repos/{REPO}/releases/latest"
SHORTCUT_ID = "chint-etm-mdm"
KEEP_VERSIONS = 2

Progress = Callable[[int, int], None]


@dataclass(frozen=True)
class Release:
    version: str
    download_url: str
    notes: str


def is_windows() -> bool:
    return sys.platform.startswith("win")


def install_root() -> Path:
    if is_windows():
        base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
        return Path(base) / APP_NAME
    base = os.environ.get("XDG_DATA_HOME") or str(Path.home() / ".local" / "share")
    return Path(base) / SHORTCUT_ID


def executable_name() -> str:
    return f"{APP_NAME}.exe" if is_windows() else APP_NAME


def version_key(version: str) -> tuple[int, ...]:
    return tuple(int(part) for part in re.findall(r"\d+", version))


def is_newer(candidate: str, current: str = __version__) -> bool:
    return version_key(candidate) > version_key(current)


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def running_installed(root: Path | None = None) -> bool:
    """True when this process is a release build started from the install folder."""
    if not is_frozen():
        return False
    root = root or install_root()
    try:
        Path(sys.executable).resolve().relative_to((root / "versions").resolve())
    except ValueError:
        return False
    return True


def load_state(root: Path) -> dict:
    try:
        data = json.loads((root / "state.json").read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def save_state(root: Path, state: dict) -> None:
    root.mkdir(parents=True, exist_ok=True)
    tmp = root / "state.json.tmp"
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, root / "state.json")


def version_exe(root: Path, version: str) -> Path:
    return root / "versions" / version / executable_name()


def previous_version(root: Path | None = None) -> str:
    root = root or install_root()
    previous = str(load_state(root).get("previous") or "")
    return previous if previous and version_exe(root, previous).exists() else ""


def _open_url(url: str, timeout: int):
    return open_url(urllib.request.Request(url, headers={"User-Agent": APP_NAME}), timeout)


def asset_matches(name: str, windows: bool) -> bool:
    name = name.lower()
    if windows:
        return name.endswith(".zip")
    return name.endswith(".tar.gz") and "linux" in name


def parse_release(data: dict, windows: bool | None = None) -> Release | None:
    windows = is_windows() if windows is None else windows
    version = str(data.get("tag_name") or "").lstrip("vV")
    for asset in data.get("assets") or []:
        if asset_matches(str(asset.get("name") or ""), windows):
            return Release(version, str(asset["browser_download_url"]), str(data.get("body") or ""))
    return None


def latest_release() -> Release | None:
    with _open_url(LATEST_RELEASE_URL, timeout=15) as response:
        return parse_release(json.loads(response.read().decode("utf-8")))


def download(url: str, target: Path, progress: Progress | None = None) -> None:
    with _open_url(url, timeout=60) as response, target.open("wb") as handle:
        total = int(response.headers.get("Content-Length") or 0)
        done = 0
        while chunk := response.read(1 << 20):
            handle.write(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)


def _extract(archive: Path, target: Path) -> None:
    if archive.name.endswith(".zip"):
        with zipfile.ZipFile(archive) as zf:
            zf.extractall(target)
    else:
        with tarfile.open(archive) as tf:
            if hasattr(tarfile, "data_filter"):
                tf.extractall(target, filter="data")
            else:
                tf.extractall(target)


def unpack_version(archive: Path, root: Path, version: str) -> Path:
    """Unpack a release archive into versions/<version> and return the executable."""
    versions = root / "versions"
    versions.mkdir(parents=True, exist_ok=True)
    staging = Path(tempfile.mkdtemp(prefix=f".{version}-", dir=versions))
    try:
        _extract(archive, staging)
        return place_version(staging, root, version)
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def place_version(unpacked: Path, root: Path, version: str) -> Path:
    """Move an unpacked app folder (or a folder holding it) to versions/<version>."""
    app_dir = unpacked / APP_NAME if (unpacked / APP_NAME).is_dir() else unpacked
    if not (app_dir / executable_name()).exists():
        raise RuntimeError(f"В архиве нет программы {executable_name()}")
    final = root / "versions" / version
    if final.exists():
        shutil.rmtree(final)
    shutil.move(str(app_dir), str(final))
    exe = final / executable_name()
    if not is_windows():
        exe.chmod(exe.stat().st_mode | 0o755)
    return exe


def desktop_dir() -> Path | None:
    if is_windows():
        return None  # the PowerShell script asks Windows for the real Desktop
    try:
        out = subprocess.run(["xdg-user-dir", "DESKTOP"], capture_output=True, text=True, timeout=5)
        path = Path(out.stdout.strip()) if out.returncode == 0 and out.stdout.strip() else None
    except (OSError, subprocess.SubprocessError):
        path = None
    path = path or Path.home() / "Desktop"
    return path if path.is_dir() else None


def linux_desktop_entry(exe: Path) -> str:
    return "\n".join(
        [
            "[Desktop Entry]",
            "Type=Application",
            f"Name={APP_NAME}",
            "Comment=Заполнение файлов для ЭТМ и РАЭК из базы CHINT",
            f'Exec="{exe}"',
            f"Path={exe.parent}",
            "Icon=x-office-spreadsheet",
            "Terminal=false",
            "Categories=Office;",
            "",
        ]
    )


def write_shortcuts(exe: Path) -> None:
    """Point the desktop and menu icons at this executable."""
    if is_windows():
        script = (
            "$s = New-Object -ComObject WScript.Shell; "
            "foreach ($d in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) { "
            f"$l = $s.CreateShortcut((Join-Path $d '{APP_NAME}.lnk')); "
            "$l.TargetPath = $env:CHINT_EXE; $l.WorkingDirectory = (Split-Path $env:CHINT_EXE); $l.Save() }"
        )
        subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
            env={**os.environ, "CHINT_EXE": str(exe)},
            check=True,
            timeout=60,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        return
    entry = linux_desktop_entry(exe)
    apps = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "applications"
    apps.mkdir(parents=True, exist_ok=True)
    (apps / f"{SHORTCUT_ID}.desktop").write_text(entry, encoding="utf-8")
    desktop = desktop_dir()
    if desktop:
        path = desktop / f"{SHORTCUT_ID}.desktop"
        path.write_text(entry, encoding="utf-8")
        path.chmod(0o755)
        subprocess.run(["gio", "set", str(path), "metadata::trusted", "true"], capture_output=True, check=False)


def activate(root: Path, version: str) -> Path:
    """Make a version current: icons point at it, the old current becomes 'previous'."""
    exe = version_exe(root, version)
    write_shortcuts(exe)
    state = load_state(root)
    current = str(state.get("current") or "")
    if current and current != version:
        state["previous"] = current
    state["current"] = version
    save_state(root, state)
    return exe


def rollback(root: Path | None = None) -> Path:
    root = root or install_root()
    previous = previous_version(root)
    if not previous:
        raise RuntimeError("Предыдущей версии нет")
    return activate(root, previous)


def cleanup_old_versions(root: Path | None = None) -> None:
    """Remove versions other than current and previous (a running one may refuse: fine)."""
    root = root or install_root()
    state = load_state(root)
    keep = {str(state.get("current") or ""), str(state.get("previous") or "")}
    versions = root / "versions"
    if not versions.is_dir():
        return
    for path in versions.iterdir():
        if path.name not in keep:
            shutil.rmtree(path, ignore_errors=True)


def install_release(release: Release, root: Path | None = None, progress: Progress | None = None) -> Path:
    root = root or install_root()
    root.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=root) as tmp:
        archive = Path(tmp) / release.download_url.rsplit("/", 1)[-1]
        download(release.download_url, archive, progress)
        unpack_version(archive, root, release.version)
    return activate(root, release.version)


def install_running_copy(root: Path | None = None) -> Path:
    """Copy this portable build (unpacked anywhere) into the install folder."""
    root = root or install_root()
    source = Path(sys.executable).resolve().parent
    staging = Path(tempfile.mkdtemp(prefix=f".{__version__}-", dir=_mkdir(root / "versions")))
    try:
        shutil.copytree(source, staging / APP_NAME)
        place_version(staging, root, __version__)
    finally:
        shutil.rmtree(staging, ignore_errors=True)
    return activate(root, __version__)


def _mkdir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def launch(exe: Path) -> None:
    """Start the program detached from this process."""
    if is_windows():
        flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
        subprocess.Popen([str(exe)], cwd=exe.parent, creationflags=flags, close_fds=True)
    else:
        subprocess.Popen([str(exe)], cwd=exe.parent, start_new_session=True, close_fds=True)


def installed_exe(root: Path | None = None) -> Path | None:
    root = root or install_root()
    current = str(load_state(root).get("current") or "")
    exe = version_exe(root, current) if current else None
    return exe if exe and exe.exists() else None


def main(argv: list[str] | None = None) -> int:
    """`python -m chint_etm_mdm.updater --install-latest`: move a source install onto release builds.

    Prints the installed executable path; exits 1 when nothing was installed.
    """
    args = argv if argv is not None else sys.argv[1:]
    if args != ["--install-latest"]:
        print("usage: python -m chint_etm_mdm.updater --install-latest", file=sys.stderr)
        return 2
    exe = installed_exe()
    if exe is None:
        try:
            release = latest_release()
        except Exception as exc:  # no internet, no release yet: keep the source install
            print(f"no release: {exc}", file=sys.stderr)
            return 1
        if release is None:
            print("no release for this system", file=sys.stderr)
            return 1
        exe = install_release(release)
    print(exe)
    return 0


if __name__ == "__main__":
    sys.exit(main())

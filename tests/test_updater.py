import io
import os
import sys
import tarfile
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from chint_etm_mdm import updater


def make_archive(folder: Path, version: str) -> Path:
    """A release archive like build_linux.sh makes: one 'CHINT ETM MDM' folder."""
    archive = folder / f"CHINT_ETM_MDM_portable_linux_{version}.tar.gz"
    with tarfile.open(archive, "w:gz") as tf:
        data = f"#!/bin/sh\necho {version}\n".encode()
        info = tarfile.TarInfo(f"{updater.APP_NAME}/{updater.APP_NAME}")
        info.size = len(data)
        info.mode = 0o755
        tf.addfile(info, io.BytesIO(data))
    return archive


@unittest.skipIf(sys.platform.startswith("win"), "Linux install layout")
class UpdaterTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        self.home = Path(self.tmp.name)
        env = {"HOME": str(self.home), "XDG_DATA_HOME": str(self.home / "share")}
        patcher = mock.patch.dict(os.environ, env)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(self.tmp.cleanup)
        (self.home / "Desktop").mkdir()
        self.root = updater.install_root()

    def install(self, version: str) -> Path:
        updater.unpack_version(make_archive(self.home, version), self.root, version)
        return updater.activate(self.root, version)

    def test_version_compare(self) -> None:
        self.assertTrue(updater.is_newer("0.10.0", "0.9.3"))
        self.assertFalse(updater.is_newer("0.2.0", "0.2.0"))
        self.assertTrue(updater.is_newer("v1.0", "0.2.0"))

    def test_picks_asset_for_system(self) -> None:
        data = {
            "tag_name": "v0.3.0",
            "assets": [
                {"name": "CHINT_ETM_MDM_portable.zip", "browser_download_url": "https://x/win.zip"},
                {"name": "CHINT_ETM_MDM_portable_linux.tar.gz", "browser_download_url": "https://x/lin.tar.gz"},
            ],
        }
        self.assertEqual("https://x/lin.tar.gz", updater.parse_release(data, windows=False).download_url)
        self.assertEqual("https://x/win.zip", updater.parse_release(data, windows=True).download_url)
        self.assertEqual("0.3.0", updater.parse_release(data, windows=False).version)
        self.assertIsNone(updater.parse_release({"tag_name": "v1", "assets": []}, windows=False))

    def test_install_points_icons_at_new_version(self) -> None:
        exe = self.install("0.2.0")
        self.assertTrue(os.access(exe, os.X_OK))
        menu_icon = self.home / "share" / "applications" / "chint-etm-mdm.desktop"
        self.assertIn(f'Exec="{exe}"', menu_icon.read_text(encoding="utf-8"))
        self.assertTrue((self.home / "Desktop" / "chint-etm-mdm.desktop").exists())
        self.assertEqual(exe, updater.installed_exe(self.root))

    def test_update_then_rollback(self) -> None:
        self.install("0.2.0")
        new_exe = self.install("0.3.0")
        self.assertEqual("0.2.0", updater.previous_version(self.root))
        self.assertEqual(new_exe, updater.installed_exe(self.root))
        old_exe = updater.rollback(self.root)
        self.assertEqual(updater.version_exe(self.root, "0.2.0"), old_exe)
        self.assertEqual("0.3.0", updater.previous_version(self.root))

    def test_keeps_only_current_and_previous(self) -> None:
        for version in ("0.2.0", "0.3.0", "0.4.0"):
            self.install(version)
        updater.cleanup_old_versions(self.root)
        self.assertEqual(["0.3.0", "0.4.0"], sorted(p.name for p in (self.root / "versions").iterdir()))

    def test_broken_archive_keeps_current_version(self) -> None:
        self.install("0.2.0")
        bad = self.home / "bad.tar.gz"
        with tarfile.open(bad, "w:gz"):
            pass
        with self.assertRaises(RuntimeError):
            updater.unpack_version(bad, self.root, "0.3.0")
        self.assertEqual(updater.version_exe(self.root, "0.2.0"), updater.installed_exe(self.root))
        self.assertEqual(["0.2.0"], [p.name for p in (self.root / "versions").iterdir()])

    def test_source_run_is_not_installed(self) -> None:
        self.assertFalse(updater.running_installed(self.root))


if __name__ == "__main__":
    unittest.main()

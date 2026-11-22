"""Build the windowed application on the current OS."""

import os
import subprocess
import sys
from pathlib import Path

root = Path(__file__).resolve().parent
icon = root / "assets" / ("app.icns" if sys.platform == "darwin" else "app.ico")
environment = os.environ.copy()
if sys.platform == "win32":
    # Resolve Qt's Windows dependencies without collecting DLLs from unrelated tools on PATH.
    windows = Path(os.environ.get("SystemRoot") or os.environ["WINDIR"])
    environment["PATH"] = os.pathsep.join(map(str, (
        windows / "System32", windows, Path(sys.executable).parent, Path(sys.base_prefix),
    )))
subprocess.run(
    [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
     "--windowed", "--onedir", "--name", "2GIS-Export",
     "--osx-bundle-identifier", "io.github.jafuart.2gis-export",
     "--icon", str(icon), "--add-data", f"{(root / 'assets').as_posix()}:assets",
     "--add-data", f"{(root / 'licenses').as_posix()}:licenses",
     "--copy-metadata", "PySide6-Essentials", "--copy-metadata", "shiboken6",
     "--collect-all", "seleniumbase", "--collect-data", "selenium", *sys.argv[1:], "app.py"],
    cwd=root, env=environment, check=True,
)

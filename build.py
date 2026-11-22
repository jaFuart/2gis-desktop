"""Build the windowed application on the current OS."""

import subprocess
import sys
from pathlib import Path

root = Path(__file__).resolve().parent
icon = root / "assets" / ("app.icns" if sys.platform == "darwin" else "app.ico")
subprocess.run(
    [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
     "--windowed", "--onedir", "--name", "2GIS-Export",
     "--osx-bundle-identifier", "io.github.jafuart.2gis-export",
     "--icon", str(icon), "--add-data", f"{(root / 'assets').as_posix()}:assets",
     "--add-data", f"{(root / 'licenses').as_posix()}:licenses",
     "--copy-metadata", "PySide6-Essentials", "--copy-metadata", "shiboken6",
     "--collect-all", "seleniumbase", "--collect-data", "selenium", *sys.argv[1:], "app.py"],
    cwd=root, check=True,
)

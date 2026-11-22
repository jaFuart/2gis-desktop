"""Build the windowed application on the current OS."""

import subprocess
import sys
from pathlib import Path

root = Path(__file__).resolve().parent
subprocess.run(
    [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean",
     "--windowed", "--onedir", "--name", "2GIS-Export",
     "--osx-bundle-identifier", "io.github.jafuart.2gis-export",
     "--collect-all", "seleniumbase", "--collect-data", "selenium", "app.py"],
    cwd=root, check=True,
)

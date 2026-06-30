# deps.py
# Ensures required third-party libraries are installed.
# - Silently hides console windows on Windows.
# - Streams pip output to a status callback when provided.
# - Uses importlib for fast dependency checking without loading modules into memory.

import sys
import subprocess
import os
import ctypes
import importlib.util
from typing import Callable, Optional

REQUIRED = [
    ("fitz", "PyMuPDF"),
    ("docx", "python-docx"),
    ("bs4", "beautifulsoup4"),
    ("xlrd", "xlrd"),
    ("sklearn", "scikit-learn"),
]

def _hide_console_window():
    try:
        kernel32 = ctypes.windll.kernel32
        user32 = ctypes.windll.user32
        get_console = kernel32.GetConsoleWindow
        show_window = user32.ShowWindow
        SW_HIDE = 0
        hwnd = get_console()
        if hwnd:
            show_window(hwnd, SW_HIDE)
    except Exception:
        pass

def _pip_install(pkg_name: str, status_callback: Optional[Callable[[str], None]] = None):
    """Install a package using pip, capturing output and hiding the console window."""
    creationflags = 0
    # Hide new console windows on Windows
    if os.name == "nt":
        creationflags = 0x08000000  # CREATE_NO_WINDOW
    cmd = [sys.executable, "-m", "pip", "install", "--upgrade", pkg_name]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, creationflags=creationflags)
        if status_callback:
            out = proc.stdout.strip()
            err = proc.stderr.strip()
            if out:
                for line in out.splitlines():
                    status_callback(line)
            if err:
                for line in err.splitlines():
                    status_callback(line)
    except Exception as e:
        if status_callback:
            status_callback(f"pip failed for {pkg_name}: {e}")

def ensure_dependencies(status_callback: Optional[Callable[[str], None]] = None):
    _hide_console_window()
    for mod_name, pkg_name in REQUIRED:
        # Check if the package exists WITHOUT loading it into memory to save startup time
        if importlib.util.find_spec(mod_name) is not None:
            if status_callback:
                status_callback(f"{pkg_name} already installed.")
        else:
            if status_callback:
                status_callback(f"Installing {pkg_name}…")
            _pip_install(pkg_name, status_callback=status_callback)
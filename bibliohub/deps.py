import importlib
import importlib.util
import subprocess
import sys

# modulo importabile -> pacchetto pip
PACKAGES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "requests": "requests",
    "bs4": "beautifulsoup4",
    "docxtpl": "docxtpl",
    "tkcalendar": "tkcalendar",
    "win32com": "pywin32",
}


def missing_packages(modules) -> list[str]:
    return [PACKAGES[m] for m in modules if importlib.util.find_spec(m) is None]


def can_install() -> bool:
    return not getattr(sys, "frozen", False)


def install(packages) -> subprocess.CompletedProcess:
    flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--quiet", *packages],
        capture_output=True, text=True, creationflags=flags,
    )
    importlib.invalidate_caches()
    return result

import sys
from pathlib import Path

def is_frozen() -> bool:
    return getattr(sys, "frozen", False)

def internal_base() -> Path:
    """
    Base path for bundled resources (_MEIPASS).
    """
    if is_frozen():
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent

def external_base() -> Path:
    """
    Base path next to the executable or script.
    """
    if is_frozen():
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent

def internal_path(relative: str) -> Path:
    """
    For files bundled via PyInstaller datas (config, chromium, etc).
    """
    return internal_base() / relative

def external_path(relative: str) -> Path:
    """
    For files next to the EXE (.env, user-editable config).
    """
    return external_base() / relative

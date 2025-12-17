# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

# 👇 single source of truth
CRITICAL_PKGS = [
    "playwright",
    "numpy",
    "pandas",
    "lxml",
    "openpyxl",
]

datas = []
binaries = []
hiddenimports = []

for pkg in CRITICAL_PKGS:
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

# Explicit Playwright internals (important)
hiddenimports += [
    "playwright.sync_api",
    "playwright._impl._api_structures",
    "playwright._impl._driver",
    "playwright._impl._connection",
]

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=[
        ('config', 'config'),
        ('chromium', 'chromium')
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="IRMAS-AUTOMATE",
    debug=False,
    strip=False,
    upx=True,
    console=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    name="IRMAS-AUTOMATE",
)

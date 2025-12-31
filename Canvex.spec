# -*- mode: python ; coding: utf-8 -*-

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect qtawesome data files (fonts) - need to try/except in case not installed
try:
    qtawesome_datas = collect_data_files('qtawesome')
    qtawesome_hiddenimports = collect_submodules('qtawesome')
except Exception:
    qtawesome_datas = []
    qtawesome_hiddenimports = []

block_cipher = None

a = Analysis(
    ['Canvex.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('splash.png', '.'),
        ('app_icon.ico', '.'),
        ('app_icon.icns', '.'),
    ] + qtawesome_datas,
    hiddenimports=[
        'qtawesome',
        'qtpy',
        'PyQt5.QtSvg',
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.common.by',
        'webdriver_manager',
        'webdriver_manager.chrome',
        'pandas',
        'openpyxl',
        'et_xmlfile',
        'xlsxwriter',
        'PIL',
        'PIL.Image',
        'requests',
        'certifi',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Canvex',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.icns',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Canvex',
)

app = BUNDLE(
    coll,
    name='Canvex.app',
    icon='app_icon.icns',
    bundle_identifier='com.canvex.app',
    info_plist={
        'CFBundleName': 'Canvex',
        'CFBundleDisplayName': 'Canvex',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '10.13',
    },
)

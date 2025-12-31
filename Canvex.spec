# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# qtawesome (safe optional)
try:
    qtawesome_datas = collect_data_files('qtawesome')
    qtawesome_hiddenimports = collect_submodules('qtawesome')
except Exception:
    qtawesome_datas = []
    qtawesome_hiddenimports = []

block_cipher = None

a = Analysis(
    ['Canvex.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('splash.png', '.'),
        ('app_icon.ico', '.'),
    ] + qtawesome_datas,
    hiddenimports=[
        *qtawesome_hiddenimports,
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
    runtime_hooks=[],
    excludes=[],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=False,      # ✅ REQUIRED for onefile
    name='Canvex',
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon='app_icon.ico',
    runtime_tmpdir=None,         # ✅ NO _internal folder
)

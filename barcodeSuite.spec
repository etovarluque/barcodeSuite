# -*- mode: python ; coding: utf-8 -*-
import shutil, os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

# openpyxl pulls in submodules dynamically; collect them all to be safe.
# et_xmlfile is an openpyxl dependency PyInstaller often fails to detect.
hidden = collect_submodules('openpyxl')
hidden += collect_submodules('et_xmlfile')
hidden += [
    'openpyxl.cell._writer',
    'et_xmlfile',
    'edlib',
    'xlsxwriter',
    'PyQt5',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.QtPrintSupport',
    'multiprocessing',
    'concurrent.futures',
]

# openpyxl ships XML/data templates it reads at runtime.
datas = collect_data_files('openpyxl')

a = Analysis(
    ['barcodeSuite.py'],
    pathex=['.'],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
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
    name='BarcodeSuite',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='BarcodeSuite',
)

# Place resources next to the .exe (same level as _internal), not inside it.
# The app resolves paths via os.path.dirname(sys.executable), so they must live here.
_dist_root = os.path.join(DISTPATH, 'BarcodeSuite')

shutil.copy(os.path.join(SPECPATH, 'icon.ico'), os.path.join(_dist_root, 'icon.ico'))

# Bundle translations only if the folder exists (optional .ts files loaded at runtime).
_tr_src = os.path.join(SPECPATH, 'translations')
if os.path.isdir(_tr_src):
    _tr_dst = os.path.join(_dist_root, 'translations')
    if os.path.exists(_tr_dst):
        shutil.rmtree(_tr_dst)
    shutil.copytree(_tr_src, _tr_dst)

# Bundle the user guide (manual/guia_uso_barcodeSuite.html) next to the .exe.
_manual_src = os.path.join(SPECPATH, 'manual')
if os.path.isdir(_manual_src):
    _manual_dst = os.path.join(_dist_root, 'manual')
    if os.path.exists(_manual_dst):
        shutil.rmtree(_manual_dst)
    shutil.copytree(_manual_src, _manual_dst)

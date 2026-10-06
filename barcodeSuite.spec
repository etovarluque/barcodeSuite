# -*- mode: python ; coding: utf-8 -*-
# Build: pyinstaller --noconfirm barcodeSuite.spec   (output: dist/BarcodeSuite/)
import shutil, os
from PyInstaller.utils.hooks import collect_submodules

hidden = collect_submodules('_utilities')
hidden += [
    'Bio.Seq', 'edlib', 'xlsxwriter',
    'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'PyQt5.QtSvg',
    'PyQt5.QtPrintSupport', 'multiprocessing', 'concurrent.futures',
]
# openpyxl is partly imported lazily; ship all of it so .xlsx reading never
# silently disappears from the frozen app.
hidden += collect_submodules('openpyxl')
if not any(h.startswith('openpyxl.') for h in hidden):
    raise SystemExit("openpyxl is not installed in the build environment")

# Pulled in only through openpyxl submodules; never used by the app.
excluded = ['PIL', 'PIL.Image', 'Pillow', 'pandas']

a = Analysis(
    ['barcodeSuite.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

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
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='BarcodeSuite',
)

# The app resolves files next to the .exe (os.path.dirname(sys.executable)).
# _profiles is deliberately NOT copied: it holds the local NCBI API key.
_dist_root = os.path.join(DISTPATH, 'BarcodeSuite')
for _folder in ('guide',):
    _dst = os.path.join(_dist_root, _folder)
    if os.path.exists(_dst):
        shutil.rmtree(_dst)
    shutil.copytree(os.path.join(SPECPATH, _folder), _dst)
for _fname in ('icon.ico', 'LICENSE', 'README.md'):
    shutil.copy(os.path.join(SPECPATH, _fname), os.path.join(_dist_root, _fname))

# -*- mode: python ; coding: utf-8 -*-

from os.path import dirname
from pathlib import Path
from PyInstaller.building.datastruct import Tree

root = Path(SPECPATH).resolve().parent.parent
payload_root = root / 'dist' / 'app' / 'Excel for Tally'
app_payload = [
    (source, dirname(destination))
    for destination, source, _typecode in Tree(str(payload_root), prefix='payload')
]

a = Analysis(
    [str(root / 'tools' / 'installer_bootstrap.py')],
    pathex=[str(root / 'src')],
    binaries=[],
    datas=[
        *app_payload,
        (str(root / 'src' / 'tally_final_upload_build' / 'license_public_key.txt'), 'tally_final_upload_build'),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Excel for Tally Setup',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

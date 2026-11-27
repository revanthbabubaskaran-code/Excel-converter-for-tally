# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

root = Path(SPECPATH).resolve().parent

a = Analysis(
    [str(root / 'tools' / 'run_app.py')],
    pathex=[str(root / 'src')],
    binaries=[],
    datas=[
        (str(root / 'src' / 'tally_final_upload_build' / 'data' / 'output_layout_template.xlsx'), 'tally_final_upload_build/data'),
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
    [],
    name='Excel for Tally',
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
    exclude_binaries=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Excel for Tally',
)

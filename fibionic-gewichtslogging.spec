# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src/fibionic_scale_app/__main__.py'],
    pathex=['src'],
    binaries=[],
    datas=[('logo', 'logo')],
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
    name='fibionic-gewichtslogging',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon='logo/fibionic_app_icon.ico',
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
app = BUNDLE(
    exe,
    name='fibionic-gewichtslogging.app',
    icon='logo/fibionic_app_icon.icns',
    bundle_identifier=None,
)

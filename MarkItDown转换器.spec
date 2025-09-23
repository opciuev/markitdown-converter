# -*- mode: python ; coding: utf-8 -*-


import os
import magika

# 获取 Magika 模型文件路径
magika_path = os.path.dirname(magika.__file__)
magika_models_path = os.path.join(magika_path, 'models')
magika_config_path = os.path.join(magika_path, 'config')

a = Analysis(
    ['markitdown_ui.py'],
    pathex=[],
    binaries=[],
    datas=[
        (magika_models_path, 'magika/models'),
        (magika_config_path, 'magika/config'),
    ],
    hiddenimports=['magika', 'markitdown'],
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
    name='MarkItDown转换器',
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

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# 收集 magika 模型文件
magika_datas = collect_data_files('magika', include_py_files=False)

a = Analysis(
    ['markitdown_ui.py'],
    pathex=[],
    binaries=[],
    datas=magika_datas,  # 添加 magika 数据文件
    hiddenimports=[
        'markitdown',
        'openpyxl',
        'PIL',
        'PIL.Image',
        'PyPDF2',
        'pdfplumber',
        'docx',
        'pptx',
        'beautifulsoup4',
        'bs4',
        'requests',
        'lxml',
        'lxml.etree',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'pdfminer',
        'pdfminer.six',
        'python_docx',
        'python_pptx',
        'charset_normalizer',
        'defusedxml',
        'markdownify',
        'mammoth',
        'xlrd',
        'et_xmlfile',
        'magika',
        'magika.magika',
        'onnxruntime',
        'onnxruntime.capi',
        'onnxruntime.capi._pybind_state',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'pandas',
        'jupyter',
        'IPython',
        'PyQt5',
        'PyQt6',
        'PyQt4',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='MarkItDownConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 改为 True 以便调试，看到错误信息
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

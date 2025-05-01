# PDFReader.spec
# Run `pyinstaller PDFReader.spec` to build

block_cipher = None

from PyInstaller.utils.hooks import collect_data_files

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=collect_data_files('pandas') + collect_data_files('pdfplumber') + collect_data_files('tkinterdnd2'),
    hiddenimports=[
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.skiplist',
        'tkinter.ttk',
        'tkinter',
        'openpyxl.styles',
        'openpyxl.formatting.rule',
    ],
    hookspath=['pyinstaller-hooks'],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PDFReader',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PDFReader'
)

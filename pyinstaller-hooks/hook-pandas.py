from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all data files and submodules for pandas
datas = collect_data_files('pandas')
hiddenimports = collect_submodules('pandas')

# Add hidden imports specific to pandas
hiddenimports.extend([
    'pandas._libs.tslibs.nattype',
    'pandas._libs.tslibs.np_datetime',
    'pandas._libs.tslibs.timestamps',
    'pandas._libs.skiplist',
])

# Optionally, collect additional data files or submodules for other packages
datas += collect_data_files('numpy')
hiddenimports.extend(collect_submodules('numpy'))

datas += collect_data_files('pdfplumber')
hiddenimports.extend(collect_submodules('pdfplumber'))

datas += collect_data_files('tkinterdnd2')
hiddenimports.extend(collect_submodules('tkinterdnd2'))

datas += collect_data_files('openpyxl')
hiddenimports.extend(collect_submodules('openpyxl'))

# You can also add any additional files manually if needed

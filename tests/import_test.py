# import_test.py

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import pandas as pd
    import pdfplumber
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill
    from openpyxl.formatting.rule import CellIsRule
    import warnings
    import threading
    import subprocess

    print("All key libraries imported successfully.")
    print("Pandas version:", pd.__version__)

except Exception as e:
    print("Import test failed:", e)
    raise

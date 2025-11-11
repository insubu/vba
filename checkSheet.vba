import ctypes
import os
import urllib.parse
from tkinter import filedialog, Tk
import win32com.client
from win32com.client import constants
import traceback

def read_ini_sheet(sheet, key: str) -> str:
    try:
        row_max = sheet.UsedRange.Rows.Count
        for row in range(1, row_max + 1):
            buf = str(sheet.Cells(row, 1).Value or "").strip()
            if buf == "":
                break
            if buf.upper() == key.upper():
                return str(sheet.Cells(row, 2).Value or "")
        return ""
    except Exception as e:
        print(f"[ReadIniSheet Error] {e}")
        return ""

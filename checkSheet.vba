import win32com.client

# Open Excel application
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Set to True if you want to see Excel

# Open the workbook
workbook = excel.Workbooks.Open(r"C:\path\to\your\file.xlsx")  # Replace with your actual path

# Access the sheet named 'iniSheet'
ini_sheet = workbook.Sheets("iniSheet")

# Define the function (already provided)
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

# Example usage
value = read_ini_sheet(ini_sheet, "ServerIP")
print(f"Value for 'ServerIP': {value}")

# Optional: close Excel
workbook.Close(SaveChanges=False)
excel.Quit()

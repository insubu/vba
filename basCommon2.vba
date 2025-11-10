'カスタムドキュメントプロパティ取得
'※プロパティの名前から値を取得
'※プロパティの設定はブックの[プロパティ]→[カスタム]にて設定可能
Public Function GetCustomDocumentProperties(book As Workbook, strName As String) As String
    On Error GoTo ErrHandler
    Dim strRet As String
    Dim objProperty As Object
    
    For Each objProperty In book.CustomDocumentProperties
        If objProperty.Name = strName Then
            strRet = objProperty.value
            Exit For
        End If
    Next

    GetCustomDocumentProperties = strRet
EndHandler:
    On Error Resume Next
    Set objProperty = Nothing
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "GetCustomDocumentProperties" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'INI形式シートの設定値を読み込む
Public Function ReadIniSheet(sheet As Worksheet, key As String) As String
    Dim lngRowMax As Long
    Dim lngRow As Long
    Dim strBuf As String
    Dim strRet As String
    
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    lngRowMax = sheet.UsedRange.Rows.Count
'***** 2005/3/31 y.yamada upd-end
    
    'シートの行数分繰り返す
    For lngRow = 1 To lngRowMax
        strBuf = Trim(sheet.Cells(lngRow, 1).value)
        If strBuf = "" Then Exit For
        
        If UCase(strBuf) = UCase(key) Then
            '該当する値を取得する
            strRet = sheet.Cells(lngRow, 2).value
            Exit For
        End If
    Next lngRow
    
    ReadIniSheet = strRet
End Function

'CSV形式シートの設定値を読み込む
Public Function ReadCsvSheet(sheet As Worksheet, key As String, Row As Long) As String
    Dim lngColMax As Long
    Dim lngCol As Long
    Dim strBuf As String
    Dim strRet As String
    
    lngColMax = sheet.UsedRange.Columns.Count
    
    'シートの列数分繰り返す
    For lngCol = 1 To lngColMax
        strBuf = Trim(sheet.Cells(2, lngCol).value)
        If strBuf = "" Then Exit For
        
        If UCase(strBuf) = UCase(key) Then
            '該当する値を取得する
            strRet = sheet.Cells(Row, lngCol).value
            Exit For
        End If
    Next lngCol
    
    ReadCsvSheet = strRet
End Function

'INI形式シートの指定セルの色を取得する
Public Function GetIniSheetColor(sheet As Worksheet, key As String) As Long
    Dim lngRowMax As Long
    Dim lngRow As Long
    Dim strBuf As String
    Dim lngRet As Long
    
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    lngRowMax = sheet.UsedRange.Rows.Count
'***** 2005/3/31 y.yamada upd-end
    
    'シートの行数分繰り返す
    For lngRow = 1 To lngRowMax
        strBuf = Trim(sheet.Cells(lngRow, 1).value)
        If strBuf = "" Then Exit For
        
        If UCase(strBuf) = UCase(key) Then
            '該当するセルの色を取得する
            lngRet = sheet.Cells(lngRow, 2).Interior.ColorIndex
            Exit For
        End If
    Next lngRow
    
    GetIniSheetColor = lngRet
End Function

'シートを削除する
Public Function DeleteSheet(book As Workbook, sheetName As String) As Boolean
    On Error GoTo ErrHandler
    
    Application.DisplayAlerts = False
    book.Worksheets(sheetName).Delete
        
    DeleteSheet = True
EndHandler:
    On Error Resume Next
    Application.DisplayAlerts = True
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "DeleteSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'シートをコピーする
Public Function CopySheet(book As Workbook, sheetName As String, newSheetName As String) As Boolean
    On Error GoTo ErrHandler
    
    Application.DisplayAlerts = False
    book.Worksheets(sheetName).Copy after:=book.Worksheets(book.Worksheets.Count)
    book.Worksheets(book.Worksheets.Count).Name = newSheetName
        
    CopySheet = True
EndHandler:
    On Error Resume Next
    Application.DisplayAlerts = True
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "CopySheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'文字列を指定バイト数になるように指定文字で埋める
Public Function FillStr(dataString As String, fillChar As String, fillByte As Integer, Optional fillRightFlag As Boolean = True) As String
    Dim strRet As String
    Dim intFillCount As Integer
    
    intFillCount = fillByte - LenB(StrConv(dataString, vbFromUnicode))
    If intFillCount >= 1 Then
        If fillRightFlag = True Then
            strRet = dataString & String(intFillCount, fillChar)
        Else
            strRet = String(intFillCount, fillChar) & dataString
        End If
'***** 2005/4/11 y.yamada ins-str
    Else
        strRet = dataString
'***** 2005/4/11 y.yamada ins-end
    End If

    FillStr = strRet
End Function

----------------------------------------------------------------
import win32com.client
from win32com.client import constants
import traceback


# =========================================================
# カスタムドキュメントプロパティ取得
# =========================================================
def get_custom_document_properties(book, name: str) -> str:
    """
    指定されたExcel Workbookオブジェクトのカスタムドキュメントプロパティから値を取得する
    """
    try:
        props = book.CustomDocumentProperties
        for prop in props:
            if prop.Name == name:
                return str(prop.Value)
        return ""
    except Exception as e:
        print(f"[GetCustomDocumentProperties Error] {e}")
        print(traceback.format_exc())
        return ""


# =========================================================
# INI形式シートの設定値を読み込む
# =========================================================
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


# =========================================================
# CSV形式シートの設定値を読み込む
# =========================================================
def read_csv_sheet(sheet, key: str, row: int) -> str:
    try:
        col_max = sheet.UsedRange.Columns.Count
        for col in range(1, col_max + 1):
            buf = str(sheet.Cells(2, col).Value or "").strip()
            if buf == "":
                break
            if buf.upper() == key.upper():
                return str(sheet.Cells(row, col).Value or "")
        return ""
    except Exception as e:
        print(f"[ReadCsvSheet Error] {e}")
        return ""


# =========================================================
# INI形式シートの指定セルの色を取得する
# =========================================================
def get_ini_sheet_color(sheet, key: str) -> int:
    try:
        row_max = sheet.UsedRange.Rows.Count
        for row in range(1, row_max + 1):
            buf = str(sheet.Cells(row, 1).Value or "").strip()
            if buf == "":
                break
            if buf.upper() == key.upper():
                return sheet.Cells(row, 2).Interior.ColorIndex
        return 0
    except Exception as e:
        print(f"[GetIniSheetColor Error] {e}")
        return 0


# =========================================================
# シートを削除する
# =========================================================
def delete_sheet(book, sheet_name: str) -> bool:
    try:
        app = book.Application
        app.DisplayAlerts = False
        book.Worksheets(sheet_name).Delete()
        return True
    except Exception as e:
        print(f"[DeleteSheet Error] {e}")
        return False
    finally:
        book.Application.DisplayAlerts = True


# =========================================================
# シートをコピーする
# =========================================================
def copy_sheet(book, sheet_name: str, new_sheet_name: str) -> bool:
    try:
        app = book.Application
        app.DisplayAlerts = False
        book.Worksheets(sheet_name).Copy(After=book.Worksheets(book.Worksheets.Count))
        new_sheet = book.Worksheets(book.Worksheets.Count)
        new_sheet.Name = new_sheet_name
        return True
    except Exception as e:
        print(f"[CopySheet Error] {e}")
        return False
    finally:
        book.Application.DisplayAlerts = True


# =========================================================
# 文字列を指定バイト数になるように指定文字で埋める
# =========================================================
def fill_str(data_string: str, fill_char: str, fill_byte: int, fill_right_flag: bool = True) -> str:
    """
    VBAのLenB(StrConv(..., vbFromUnicode))に相当するバイト数で埋める。
    PythonではShift_JISで近似。
    """
    try:
        byte_len = len(data_string.encode("shift_jis", errors="ignore"))
        fill_count = fill_byte - byte_len
        if fill_count >= 1:
            if fill_right_flag:
                return data_string + (fill_char * fill_count)
            else:
                return (fill_char * fill_count) + data_string
        else:
            return data_string
    except Exception as e:
        print(f"[FillStr Error] {e}")
        return data_string


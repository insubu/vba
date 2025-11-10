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

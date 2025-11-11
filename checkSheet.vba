'チェック処理メイン
Private Function CheckMain(blnOpenFile As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim strFilter As String
    Dim strOpenFilePath As String
    Dim strWorkFilePath As String
    Dim i As Integer
    Dim varCol(255) As Variant
    Dim blnSepTab As Boolean
    Dim blnSepSemicolon As Boolean
    Dim blnSepComma As Boolean
    Dim blnSepSpace As Boolean
    Dim blnSepOther As Boolean
    Dim strSepOther As String
    
    '全設定シートを読み込む
    If ReadAllSheet = False Then GoTo EndHandler

    If blnOpenFile = True Then
        'ファイル選択ダイアログを表示する
        strFilter = "テキストファイル(*.csv,*.txt)" & vbNullChar & "*.csv;*.txt" & vbNullChar
        strOpenFilePath = OpenFileDialog(ThisWorkbook.path, "ファイルを選択してください", strFilter)
        If strOpenFilePath = "" Then
            GoTo EndHandler
        End If
        
        '選択されたファイルを拡張子"txt"のファイルとして作業ディレクトリへコピーする
        '※拡張子"csv"のタブ区切りファイルがExcelの標準機能で読み込めないため
        If objFso.FolderExists(ThisWorkbook.path & "\" & DIR_WORK) = False Then
            '作業ディレクトリが存在しない場合は作成する
            Call objFso.CreateFolder(ThisWorkbook.path & "\" & DIR_WORK)
        End If
        strWorkFilePath = ThisWorkbook.path & "\" & DIR_WORK & "\" & objFso.GetBaseName(strOpenFilePath) & ".txt"
        Call objFso.CopyFile(strOpenFilePath, strWorkFilePath, True)
        
        'セパレータを決定する
        Select Case MainInfo.TextSeparator
        Case STR_TAB
            blnSepTab = True
        Case ";"
            blnSepSemicolon = True
        Case ","
            blnSepComma = True
        Case " "
            blnSepSpace = True
        Case Else
            'その他の場合はセパレータを指定する
            blnSepOther = True
            strSepOther = MainInfo.TextSeparator
        End Select
        
        'ファイルをテキスト形式でオープンする
        For i = 0 To UBound(varCol)
            varCol(i) = Array(i + 1, 2)
        Next i
        Workbooks.OpenText fileName:=strWorkFilePath, StartRow:=1, DataType:=xlDelimited, _
                           TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, _
                           Tab:=blnSepTab, Semicolon:=blnSepSemicolon, Comma:=blnSepComma, Space:=blnSepSpace, _
                           Other:=blnSepOther, OtherChar:=strSepOther, FieldInfo:=varCol
    End If
    
    'シート上の属性値をチェックする
    If CheckSheet = False Then GoTo EndHandler
    'チェック処理後の属性情報をファイルに保存する
    If SaveResultToFile = False Then GoTo EndHandler
    '正常終了メッセージ
    Call OutputMsg(MSG_202, MODE_DLG, "", vbInformation, APP_TITLE)

    CheckMain = True
EndHandler:
    On Error Resume Next
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "CheckMain" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function


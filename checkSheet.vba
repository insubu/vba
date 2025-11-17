'チェック処理後の属性情報をファイルに保存する
Public Function SaveResultToFile() As Boolean
    On Error GoTo ErrHandler
    Dim intFileNum As Integer
    Dim strDirPath As String
    Dim strFileName As String
    Dim lngRow As Long
    Dim lngRowMin As Long
    Dim lngRowMax As Long
    Dim lngCol As Long
    Dim lngColMax As Long
    Dim strHeader As String
    Dim strData As String
    Dim strBuf As String
    Dim lngAttributeInfoIndex As Long

    '保存先ディレクトリ
    strDirPath = MainInfo.SaveDirPath
    If strDirPath = "" Then
        '未指定の場合は処理対象ブックと同じ位置
        strDirPath = bokTarget.path
    End If
    '保存ファイル名
    strFileName = objFso.GetBaseName(bokTarget.FullName) & "_" & Format(Now, "YYYYMMDDhhmmss") & "." & MainInfo.SaveExtension

    '書き込みモードでファイルをオープンする
    intFileNum = FreeFile()
    Open strDirPath & "\" & strFileName For Output As #intFileNum
    
    '最終行列を取得する
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = shtTarget.Cells(Rows.Count, MainInfo.OriginCell.Col).End(xlUp).Row
'    lngColMax = shtTarget.Cells(MainInfo.OriginCell.Row, Columns.Count).End(xlToLeft).Column
    lngRowMax = shtTarget.UsedRange.Rows.Count
    lngColMax = shtTarget.UsedRange.Columns.Count
'***** 2005/3/31 y.yamada upd-end
    
    If MainInfo.OriginCell.AddHeader = False And MainInfo.SaveMode <> enumSaveMode.Fixed Then
        'ヘッダ追加モードでなく、固定長でない場合はヘッダを出力する
        lngRowMin = MainInfo.OriginCell.Row
    Else
        'それ以外の場合はヘッダを出力しない
        lngRowMin = MainInfo.OriginCell.Row + 1
    End If
    
    '行数分繰り返す
    For lngRow = lngRowMin To lngRowMax
        strBuf = ""
        '列数分繰り返す
        For lngCol = MainInfo.OriginCell.Col To lngColMax - 1
            'ヘッダとデータを取得する
            strHeader = shtTarget.Cells(MainInfo.OriginCell.Row, lngCol).value
            strData = shtTarget.Cells(lngRow, lngCol).value
            
            If lngRow <> MainInfo.OriginCell.Row Then
                '属性名をキーとして属性情報におけるインデックスを取得する
                lngAttributeInfoIndex = GetAttributeInfoIndex(strHeader)
                If lngAttributeInfoIndex = -1 Then
                    '属性未定義エラー
                    Call OutputMsg(MSG_104, MODE_DLG, strHeader, vbExclamation, APP_TITLE)
                    GoTo EndHandler
                End If
                
                Select Case AttributeInfo(lngAttributeInfoIndex).AttrType
                Case enumAttributeType.Alphanumeric, enumAttributeType.Date, enumAttributeType.Narrow, enumAttributeType.NarrowKana, enumAttributeType.Wide
                    '数値型以外の場合はデータをダブルコートで囲む（固定長を除く）
                    If MainInfo.SaveMode <> enumSaveMode.Fixed Then
                        strData = PutDQ(strData)
                    End If
                End Select
            End If
            If lngCol <> MainInfo.OriginCell.Col Then
                'ファイル保存方法に応じたセパレータで連結していく
                Select Case MainInfo.SaveMode
                Case enumSaveMode.Csv, enumSaveMode.TextComma
                    strBuf = strBuf & ","
                Case enumSaveMode.TextTab
                    strBuf = strBuf & vbTab
                End Select
            End If
            strBuf = strBuf & strData
        Next lngCol
        '１行分書き込む
        Print #intFileNum, strBuf
    Next lngRow
    
    SaveResultToFile = True
EndHandler:
    On Error Resume Next
    Close intFileNum
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "SaveResultToFile" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    GoTo EndHandler
End Function


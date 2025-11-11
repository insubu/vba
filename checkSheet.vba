'シート上の属性値をチェックする
Private Function CheckSheet() As Boolean
    On Error GoTo ErrHandler
    Dim lngRow As Long
    Dim lngRowMax As Long
    Dim lngCol As Long
    Dim lngColMax As Long
    Dim lngColResult As Long
    Dim strHeader As String
    Dim strData As String
    Dim strBuf As String
    Dim strErrMsg As String
    Dim lngAttributeInfoIndex As Long
    Dim blnErrFlag As Boolean
    Dim lngDataNumber As Long
    
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    '作業シートを作成する（以降このシートを処理対象とする）
    If MakeWorkSheet = False Then GoTo EndHandler
    
    '処理対象シートの最終行列を取得する
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = shtTarget.Cells(Rows.Count, MainInfo.OriginCell.Col).End(xlUp).Row
'    lngColMax = shtTarget.Cells(MainInfo.OriginCell.Row, Columns.Count).End(xlToLeft).Column
    lngRowMax = shtTarget.UsedRange.Rows.Count
    lngColMax = shtTarget.UsedRange.Columns.Count
'***** 2005/3/31 y.yamada upd-end
    
    '最終列に処理結果列を作成する
    lngColResult = lngColMax + 1
    shtTarget.Cells(MainInfo.OriginCell.Row, lngColResult).value = MainInfo.ResultHeadrName
    
    '行数分繰り返す
    For lngRow = MainInfo.OriginCell.Row + 1 To lngRowMax
        '処理経過を表示する
        lngDataNumber = lngDataNumber + 1
        Application.StatusBar = APP_TITLE & " " & "処理中です...[" & CStr(lngDataNumber) & "/" & CStr(lngRowMax - MainInfo.OriginCell.Row) & "件]"
        
        '列数分繰り返す
        For lngCol = MainInfo.OriginCell.Col To lngColMax
            'ヘッダとデータを取得する
            strHeader = shtTarget.Cells(MainInfo.OriginCell.Row, lngCol).value
            strData = shtTarget.Cells(lngRow, lngCol).value
            
            '属性名をキーとして属性情報におけるインデックスを取得する
            lngAttributeInfoIndex = GetAttributeInfoIndex(strHeader)
            If lngAttributeInfoIndex = -1 Then
                '属性未定義エラー
                Call OutputMsg(MSG_104, MODE_DLG, strHeader, vbExclamation, APP_TITLE)
                GoTo EndHandler
            End If
            
            '値を編集する
            If EditValue(strData, lngAttributeInfoIndex, strErrMsg, strBuf) = False Then
                '編集エラー
                '　エラーセルに色を付ける
                shtTarget.Cells(lngRow, lngCol).Interior.ColorIndex = MainInfo.ErrCellColor
                '　処理結果列にエラーメッセージを表示する
                shtTarget.Cells(lngRow, lngColResult).value = RESULT_NG & " [" & strHeader & ":" & strErrMsg & "]"
                blnErrFlag = True   'エラー存在フラグON
                '　次の行へ
                Exit For
            End If
            If strData <> strBuf Then
                '編集されたセルに色を付ける
                shtTarget.Cells(lngRow, lngCol).Interior.ColorIndex = MainInfo.EditCellColor
                shtTarget.Cells(lngRow, lngCol).value = strBuf
            End If
        Next lngCol
        If lngCol = lngColMax + 1 Then
            '全列処理できた場合は正常終了
            shtTarget.Cells(lngRow, lngColResult).value = RESULT_OK
        End If
    Next lngRow
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    If blnErrFlag = True Then
        '１つでもエラーが存在する場合は確認ダイアログを表示する
        Call OutputMsg(MSG_201, MODE_DLG, "", vbExclamation, APP_TITLE)
        GoTo EndHandler
    End If
    
    CheckSheet = True
EndHandler:
    On Error Resume Next
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "CheckSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function
    

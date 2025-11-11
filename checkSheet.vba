'作業シートを作成する
Private Function MakeWorkSheet() As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    Dim lngCol As Long
    Dim lngColMax As Long
    Dim strBuf As String

    If Application.Workbooks.Count = 1 Then
        '開いているブックが1つしかない場合はエラー
        Call OutputMsg(MSG_101, MODE_DLG, "", vbExclamation, APP_TITLE)
        GoTo EndHandler
    End If
    If ActiveWorkbook.Name = ThisWorkbook.Name Then
        'アクティブなブックが自分自身の場合はエラー
        Call OutputMsg(MSG_102, MODE_DLG, "", vbExclamation, APP_TITLE)
        GoTo EndHandler
    End If
    '処理対象ブックをセットする
    Set bokTarget = ActiveWorkbook
    
    If bokTarget.path = "" And MainInfo.SaveDirPath = "" Then
        'ブックが一時ファイルでファイル保存先が未指定の場合はエラー
        Call OutputMsg(MSG_105, MODE_DLG, "", vbExclamation, APP_TITLE)
        GoTo EndHandler
    End If
        
    '処理対象シートをセットする（とりあえず）
    Set shtTarget = bokTarget.ActiveSheet
    
    'シート数分繰り返す
    For i = 1 To bokTarget.Worksheets.Count
        If bokTarget.Worksheets(i).Name = shtTarget.Name & "_" & WORK_SHEET_TAG Then
            '既に作業シートが存在する場合は再実行確認ダイアログを表示する
            If OutputMsg(MSG_103, MODE_DLG, "", vbQuestion + vbOKCancel, APP_TITLE) = vbCancel Then
                '処理中止
                GoTo EndHandler
            End If
            '作業シートを削除する
            If DeleteSheet(bokTarget, bokTarget.Worksheets(i).Name) = False Then GoTo EndHandler
            Exit For
        End If
    Next i
    
    '処理対象シートを作業シートとしてコピーする（以降このシートを処理対象とする）
    If CopySheet(bokTarget, shtTarget.Name, shtTarget.Name & "_" & WORK_SHEET_TAG) = False Then GoTo EndHandler
    Set shtTarget = bokTarget.Worksheets(bokTarget.Worksheets.Count)
        
    'ヘッダ行追加
    If MainInfo.OriginCell.AddHeader = True Then
        'データ開始行の位置に空白行を挿入する
        shtTarget.Rows(MainInfo.OriginCell.Row).Insert
        '属性情報数分繰り返す
        For i = 0 To AttributeInfoCount - 1
            '属性名をヘッダ行の該当する位置に記述する
            shtTarget.Cells(MainInfo.OriginCell.Row, MainInfo.OriginCell.Col + AttributeInfo(i).ColPos - 1).value = AttributeInfo(i).AttrName
        Next i
    End If
    
    '余白削除
    If MainInfo.OriginCell.DeleteUpperRow = True Then
        If MainInfo.OriginCell.Row > 1 Then
            'データ開始行より上の行を削除する（以降１行目をデータ開始行とする）
            shtTarget.Range(shtTarget.Rows(1), shtTarget.Rows(MainInfo.OriginCell.Row - 1)).Delete
            MainInfo.OriginCell.Row = 1
        End If
        If MainInfo.OriginCell.Col > 1 Then
            'データ開始列より左の列を削除する（以降１列目をデータ開始列とする）
            shtTarget.Range(shtTarget.Columns(1), shtTarget.Columns(MainInfo.OriginCell.Col - 1)).Delete
            MainInfo.OriginCell.Col = 1
        End If
    End If
    
    'ヘッダの改行・スペース削除
'***** 2005/3/31 y.yamada upd-str
'    lngColMax = shtTarget.Cells(MainInfo.OriginCell.Row, Columns.Count).End(xlToLeft).Column
    lngColMax = shtTarget.UsedRange.Columns.Count
'***** 2005/3/31 y.yamada upd-end
    '列数分繰り返す
    For lngCol = MainInfo.OriginCell.Col To lngColMax
        strBuf = shtTarget.Cells(MainInfo.OriginCell.Row, lngCol).value
        '改行を削除する
        If MainInfo.TrimHeaderCrLf = True Then
            strBuf = Replace(strBuf, vbCr, "")
            strBuf = Replace(strBuf, vbLf, "")
        End If
        'スペースを削除する
        Select Case MainInfo.TrimHeaderSpace
        Case enumTrimSpaceMode.TrimAll
            strBuf = Replace(strBuf, " ", "")
        Case enumTrimSpaceMode.TrimBoth
            strBuf = Trim(strBuf)
        Case enumTrimSpaceMode.TrimLeft
            strBuf = LTrim(strBuf)
        Case enumTrimSpaceMode.TrimRight
            strBuf = RTrim(strBuf)
        End Select
        shtTarget.Cells(MainInfo.OriginCell.Row, lngCol).value = strBuf
    Next lngCol
    
    MakeWorkSheet = True
EndHandler:
    On Error Resume Next
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "MakeWorkSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function


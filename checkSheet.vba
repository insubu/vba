'メインシートを読み込む
Public Function ReadMainSheet() As Boolean
    On Error GoTo ErrHandler
    Dim MainInfo_clear As MAIN_SHEET
    Dim strKey As String
    Dim strBuf As String
    Dim strWork As String
    Dim strCode As String
    Dim PeriodCode As PERIOD_CODE
        
    '初期化
    MainInfo = MainInfo_clear
        
    '設定内容をチェックする
    '　ファイル保存先
    strKey = "ファイル保存先"
    strBuf = ReadIniSheet(shtMain, strKey)
    If strBuf <> "" Then
        If objFso.FolderExists(strBuf) = False Then
            GoTo IniErrHandler
        End If
    End If
    MainInfo.SaveDirPath = strBuf
    '　ファイル保存方法
    strKey = "ファイル保存方法"
    strWork = ReadIniSheet(shtMain, strKey)
    '　＜保存モード＞
    strBuf = GetOneData(strWork, ":")
    Select Case StrConv(UCase(strBuf), vbNarrow)
    Case "CSV"
        MainInfo.SaveMode = enumSaveMode.Csv
    Case "TEXT(ﾀﾌﾞ)"
        MainInfo.SaveMode = enumSaveMode.TextTab
    Case "TEXT(ｶﾝﾏ)"
        MainInfo.SaveMode = enumSaveMode.TextComma
    Case "固定長"
        MainInfo.SaveMode = enumSaveMode.Fixed
    Case Else
        GoTo IniErrHandler
    End Select
    '　＜保存ファイル拡張子＞
    strBuf = GetOneData(strWork, ":")
    If strBuf = "" Then
        GoTo IniErrHandler
    End If
    MainInfo.SaveExtension = strBuf
    '　開始セル
    strKey = "開始セル"
    strWork = ReadIniSheet(shtMain, strKey)
    '　＜開始行＞
    strBuf = GetOneData(strWork, ",")
    If strBuf = "" Or IsNumeric(strBuf) = False Then
        GoTo IniErrHandler
    End If
    If CLng(strBuf) < 1 Or CLng(strBuf) > 65535 Then
        GoTo IniErrHandler
    End If
    MainInfo.OriginCell.Row = CLng(strBuf)
    '　＜開始列＞
    strBuf = GetOneData(strWork, ":")
    If strBuf = "" Or IsNumeric(strBuf) = False Then
        GoTo IniErrHandler
    End If
    If CLng(strBuf) < 1 Or CLng(strBuf) > 255 Then
        GoTo IniErrHandler
    End If
    MainInfo.OriginCell.Col = CLng(strBuf)
    '　＜開始行より上の行を削除＞
    strBuf = GetOneData(strWork, ":")
    Select Case UCase(strBuf)
    Case "Y"
        MainInfo.OriginCell.DeleteUpperRow = True
    Case "N"
        MainInfo.OriginCell.DeleteUpperRow = False
    Case Else
        GoTo IniErrHandler
    End Select
    '　＜ヘッダ追加＞
    strBuf = GetOneData(strWork, ":")
    Select Case UCase(strBuf)
    Case "ADD"
        MainInfo.OriginCell.AddHeader = True
    Case "HEADER"
        MainInfo.OriginCell.AddHeader = False
    Case Else
        GoTo IniErrHandler
    End Select
    '　セパレータ
    strKey = "セパレータ"
    strBuf = ReadIniSheet(shtMain, strKey)
    If strBuf = "" Then
        GoTo IniErrHandler
    End If
    If UCase(strBuf) <> STR_TAB And Len(strBuf) <> 1 Then
        GoTo IniErrHandler
    End If
    MainInfo.TextSeparator = strBuf
    '　エラー背景色
    strKey = "エラー背景色"
    MainInfo.ErrCellColor = GetIniSheetColor(shtMain, strKey)
    '　編集済み背景色
    strKey = "編集済背景色"
    MainInfo.EditCellColor = GetIniSheetColor(shtMain, strKey)
    '　処理結果
    strKey = "処理結果"
    strBuf = ReadIniSheet(shtMain, strKey)
    If strBuf = "" Then
        GoTo IniErrHandler
    End If
    MainInfo.ResultHeadrName = strBuf
    '　ヘッダ行スペース削除
    strKey = "ヘッダー行スペース削除"
    strBuf = ReadIniSheet(shtMain, strKey)
    Select Case strBuf
    Case ""
        MainInfo.TrimHeaderSpace = enumTrimSpaceMode.Non
    Case "全て"
        MainInfo.TrimHeaderSpace = enumTrimSpaceMode.TrimAll
    Case "前方"
        MainInfo.TrimHeaderSpace = enumTrimSpaceMode.TrimLeft
    Case "後方"
        MainInfo.TrimHeaderSpace = enumTrimSpaceMode.TrimRight
    Case "両端"
        MainInfo.TrimHeaderSpace = enumTrimSpaceMode.TrimBoth
    Case Else
        GoTo IniErrHandler
    End Select
    '　ヘッダ行改行削除
    strKey = "ヘッダー行改行削除"
    strBuf = ReadIniSheet(shtMain, strKey)
    Select Case UCase(strBuf)
    Case "Y"
        MainInfo.TrimHeaderCrLf = True
    Case "N"
        MainInfo.TrimHeaderCrLf = False
    Case Else
        GoTo IniErrHandler
    End Select
    '　句点コード
    strKey = "句点コード"
    strWork = ReadIniSheet(shtMain, strKey)
    Do While strWork <> ""
        strCode = GetOneData(strWork, ",")
        '＜開始＞
        strBuf = GetOneData(strCode, ":")
        If strBuf = "" Then
            GoTo IniErrHandler
        End If
        PeriodCode.From = strBuf
        '＜終了＞
        strBuf = GetOneData(strCode, ":")
        If strBuf = "" Then
            '未指定の場合は＜開始＞と同じ値をセットする
            strBuf = PeriodCode.From
        End If
        PeriodCode.To = strBuf
        '情報を配列に追加する
        ReDim Preserve MainInfo.PeriodCode(MainInfo.PeriodCodeCount)
        MainInfo.PeriodCode(MainInfo.PeriodCodeCount) = PeriodCode
        MainInfo.PeriodCodeCount = MainInfo.PeriodCodeCount + 1
    Loop

    ReadMainSheet = True
EndHandler:
    On Error Resume Next
    Exit Function
IniErrHandler:      '設定エラー
    'シートをアクティブにする
    Call ShowIniSheet(shtMain)
    Call OutputMsg(MSG_001, MODE_ALL, shtMain.Name & "#" & strKey, vbCritical, APP_TITLE)
    GoTo EndHandler
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "ReadMainSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function


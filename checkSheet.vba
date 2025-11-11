Option Explicit

Private Const mModuleName As String = "basIniSheet"

'ファイル保存方法
Public Enum enumSaveMode
    Csv                 'CSV
    TextTab             'TEXT（タブ）
    TextComma           'TEXT（カンマ）
    Fixed               '固定長
End Enum
'属性の型
Public Enum enumAttributeType
    Non                 'なし
    Narrow              '半角
    Wide                '全角
    Alphanumeric        '英数字
    NarrowKana          '半角カナ
    IntegerNumber       '整数
    SmallNumber         '小数
    Date                '日付
End Enum
'文字の型
Public Enum enumLetterType
    Non                 'なし
    Capital             '大文字
    Small               '小文字
End Enum
'バイト数加工
Public Enum enumByteEditMode
    Non                 'なし
    Fixed               '固定
    Complete            '補完
    Max                 '最大
End Enum
'スペース削除
Public Enum enumTrimSpaceMode
    Non                 'なし
    TrimAll             '全て
    TrimLeft            '前方
    TrimRight           '後方
    TrimBoth            '両端
End Enum
'置換
Public Enum enumReplaceMode
    Complete            '完全一致
    Partial             '部分一致
End Enum
'開始セル
Public Type ORIGIN_CELL
    Row As Long                 '開始行
    Col As Long                 '開始列
    DeleteUpperRow As Boolean   '開始行より上の行を削除
    AddHeader As Boolean        'ヘッダ追加
End Type
'句点コード
Public Type PERIOD_CODE
    From As String              '開始
    To As String                '終了
End Type

'メインシート
Public Type MAIN_SHEET
    SaveDirPath As String                   'ファイル保存先
    SaveMode As enumSaveMode                'ファイル保存方法
    SaveExtension As String                 '保存ファイル拡張子
    OriginCell As ORIGIN_CELL               '開始セル
    TextSeparator As String                 'セパレータ
    ErrCellColor As Long                    'エラー背景色
    EditCellColor As Long                   '編集済み背景色
    ResultHeadrName As String               '処理結果
    TrimHeaderSpace As enumTrimSpaceMode    'ヘッダ行スペース削除
    TrimHeaderCrLf As Boolean               'ヘッダ行改行削除
    PeriodCode() As PERIOD_CODE             '句点コード
    PeriodCodeCount As Integer              '句点コードの件数
End Type
Public MainInfo As MAIN_SHEET

'属性シート
Public Type ATTRIBUTE_SHEET
    AttrName As String                      '属性名
    ColPos As Integer                       '属性位置
    Indispensable As Boolean                '必須
    AttrType As enumAttributeType           '属性の型
    LetterType As enumLetterType            '文字の型
    ByteSize_Left As Integer                'バイト数（整数部）
    ByteSize_Right As Integer               'バイト数（小数部）
    ByteEditMode As enumByteEditMode        'バイト数加工
    TrimSpace As enumTrimSpaceMode          'スペース削除
    TrimCrLf As Boolean                     '改行削除
    DateFormat_In As String                 '日付フォーマット（入力）
    DateFormat_Out As String                '日付フォーマット（出力）
    CompleteChar As String                  '補完文字
End Type
Public AttributeInfo() As ATTRIBUTE_SHEET
Public AttributeInfoCount As Long

'置換シート
Public Type REPLACE_SHEET
    KeyString As String                     '変換前
    ReplaceString As String                 '変換後
    ReplaceMode As enumReplaceMode          '変換モード
End Type
Public ReplaceInfo() As REPLACE_SHEET
Public ReplaceInfoCount As Long

'全設定シートを読み込む
Public Function ReadAllSheet() As Boolean
    
    'メインシートを読み込む
    If ReadMainSheet = False Then GoTo EndHandler
    '属性シートを読み込む
    If ReadAttributeSheet = False Then GoTo EndHandler
    '置換シートを読み込む
    If ReadReplaceSheet = False Then GoTo EndHandler

    ReadAllSheet = True
EndHandler:
End Function

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

'属性シートを読み込む
Public Function ReadAttributeSheet() As Boolean
    On Error GoTo ErrHandler
    Dim strKey As String
    Dim strBuf As String
    Dim lngRow As Long
    Dim lngRowMax As Long
        
    '初期化
    Erase AttributeInfo
    AttributeInfoCount = 0
        
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = shtAttribute.Cells(Rows.Count, 1).End(xlUp).Row
    lngRowMax = shtAttribute.UsedRange.Rows.Count
'***** 2005/3/31 y.yamada upd-end
    
    'シートの行数分繰り返す
    For lngRow = 3 To lngRowMax
        ReDim Preserve AttributeInfo(AttributeInfoCount)
        
        '設定内容をチェックする
        '　属性名
        strKey = "属性名"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        If strBuf = "" Then
'***** 2005/3/31 y.yamada upd-str
'            GoTo IniErrHandler
            '未入力の場合は有効行でないと判断する
            Exit For
'***** 2005/3/31 y.yamada upd-end
        End If
        AttributeInfo(AttributeInfoCount).AttrName = strBuf
        '　属性位置
        strKey = "属性位置"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        If MainInfo.OriginCell.AddHeader = True Then
            'ヘッダ追加モードの場合のみ
            If strBuf = "" Then
                GoTo IniErrHandler
            End If
            If IsNumeric(strBuf) = False Then
                GoTo IniErrHandler
            End If
            AttributeInfo(AttributeInfoCount).ColPos = CInt(strBuf)
        End If
        '　必須
        strKey = "必須"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        Select Case UCase(strBuf)
        Case "Y"
            AttributeInfo(AttributeInfoCount).Indispensable = True
        Case "N"
            AttributeInfo(AttributeInfoCount).Indispensable = False
        Case Else
            GoTo IniErrHandler
        End Select
        '　属性の型
        strKey = "型"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        Select Case strBuf
        Case ""
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.Non
        Case "半角"
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.Narrow
        Case "全角"
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.Wide
        Case "英数字"
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.Alphanumeric
        Case "半角カナ"
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.NarrowKana
        Case "整数"
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.IntegerNumber
        Case "小数"
            AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.SmallNumber
        Case Else
            If strBuf Like "日付:*" Then
                AttributeInfo(AttributeInfoCount).AttrType = enumAttributeType.Date
                AttributeInfo(AttributeInfoCount).DateFormat_In = Mid(strBuf, InStr(strBuf, ":") + 1)
            Else
                GoTo IniErrHandler
            End If
        End Select
        '　文字の型
        strKey = "大文字/小文字"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        Select Case strBuf
        Case ""
            AttributeInfo(AttributeInfoCount).LetterType = enumLetterType.Non
        Case "大文字"
            AttributeInfo(AttributeInfoCount).LetterType = enumLetterType.Capital
        Case "小文字"
            AttributeInfo(AttributeInfoCount).LetterType = enumLetterType.Small
        Case Else
            GoTo IniErrHandler
        End Select
        '　バイト数
        strKey = "バイト数"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        If strBuf <> "" Then
            If IsNumeric(strBuf) = False Then
                GoTo IniErrHandler
            End If
            If CInt(strBuf) < 1 Then
                GoTo IniErrHandler
            End If
            If InStr(strBuf, ".") <> 0 Then
                AttributeInfo(AttributeInfoCount).ByteSize_Left = CInt(Left(strBuf, InStr(strBuf, ".") - 1))
                AttributeInfo(AttributeInfoCount).ByteSize_Right = CInt(Mid(strBuf, InStr(strBuf, ".") + 1))
            Else
                AttributeInfo(AttributeInfoCount).ByteSize_Left = CInt(strBuf)
            End If
        End If
        '　バイト数加工
        strKey = "バイト数加工"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        Select Case strBuf
        Case ""
            AttributeInfo(AttributeInfoCount).ByteEditMode = enumByteEditMode.Non
        Case "固定"
            AttributeInfo(AttributeInfoCount).ByteEditMode = enumByteEditMode.Fixed
        Case "最大"
            AttributeInfo(AttributeInfoCount).ByteEditMode = enumByteEditMode.Max
        Case Else
            If strBuf Like "補完:*" Then
                AttributeInfo(AttributeInfoCount).ByteEditMode = enumByteEditMode.Complete
                Select Case AttributeInfo(AttributeInfoCount).AttrType
                Case enumAttributeType.Date
                    AttributeInfo(AttributeInfoCount).DateFormat_Out = Mid(strBuf, InStr(strBuf, ":") + 1)
                Case enumAttributeType.IntegerNumber, enumAttributeType.SmallNumber
                    Select Case Mid(strBuf, InStr(strBuf, ":") + 1)
                    Case "0", " "
                        AttributeInfo(AttributeInfoCount).CompleteChar = Mid(strBuf, InStr(strBuf, ":") + 1)
                    Case Else
                        GoTo IniErrHandler
                    End Select
                Case Else
                    If Len(Mid(strBuf, InStr(strBuf, ":") + 1)) > 1 Then
                        GoTo IniErrHandler
                    End If
                    AttributeInfo(AttributeInfoCount).CompleteChar = Mid(strBuf, InStr(strBuf, ":") + 1)
                End Select
            Else
                GoTo IniErrHandler
            End If
        End Select
        '　スペース削除
        strKey = "スペース削除"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        Select Case strBuf
        Case ""
            AttributeInfo(AttributeInfoCount).TrimSpace = enumTrimSpaceMode.Non
        Case "全て"
            AttributeInfo(AttributeInfoCount).TrimSpace = enumTrimSpaceMode.TrimAll
        Case "前方"
            AttributeInfo(AttributeInfoCount).TrimSpace = enumTrimSpaceMode.TrimLeft
        Case "後方"
            AttributeInfo(AttributeInfoCount).TrimSpace = enumTrimSpaceMode.TrimRight
        Case "両端"
            AttributeInfo(AttributeInfoCount).TrimSpace = enumTrimSpaceMode.TrimBoth
        Case Else
            GoTo IniErrHandler
        End Select
        '　改行削除
        strKey = "改行削除"
        strBuf = ReadCsvSheet(shtAttribute, strKey, lngRow)
        Select Case UCase(strBuf)
        Case "Y"
            AttributeInfo(AttributeInfoCount).TrimCrLf = True
        Case "N"
            AttributeInfo(AttributeInfoCount).TrimCrLf = False
        Case Else
            GoTo IniErrHandler
        End Select
        
        AttributeInfoCount = AttributeInfoCount + 1
    Next lngRow
    
    ReadAttributeSheet = True
EndHandler:
    On Error Resume Next
    Exit Function
IniErrHandler:      '設定エラー
    'シートをアクティブにする
    Call ShowIniSheet(shtAttribute)
    Call OutputMsg(MSG_002, MODE_ALL, shtAttribute.Name & "#" & strKey & "#" & CStr(lngRow), vbCritical, APP_TITLE)
    GoTo EndHandler
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "ReadAttributeSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'置換シートを読み込む
Public Function ReadReplaceSheet() As Boolean
    On Error GoTo ErrHandler
    Dim strBuf As String
    Dim strKey As String
    Dim lngRow As Long
    Dim lngRowMax As Long
    Dim i As Long
        
    '初期化
    Erase ReplaceInfo
    ReplaceInfoCount = 0
        
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = shtReplace.Cells(Rows.Count, 1).End(xlUp).Row
    lngRowMax = shtReplace.UsedRange.Rows.Count
'***** 2005/3/31 y.yamada upd-end
    
    'シートの行数分繰り返す
    For lngRow = 3 To lngRowMax
        ReDim Preserve ReplaceInfo(ReplaceInfoCount)
        
        '設定内容をチェックする
        '　変換前
        strKey = "変換前"
        strBuf = ReadCsvSheet(shtReplace, strKey, lngRow)
        If strBuf = "" Then
'***** 2005/3/31 y.yamada upd-str
'            GoTo IniErrHandler
            '未入力の場合は有効行でないと判断する
            Exit For
'***** 2005/3/31 y.yamada upd-end
        End If
'***** 2005/3/31 y.yamada del-str
'        For i = 0 To ReplaceInfoCount - 1
'            If InStr(ReplaceInfo(i).KeyString, strBuf) <> 0 Or InStr(strBuf, ReplaceInfo(i).KeyString) <> 0 Then
'                GoTo IniErrHandler
'            End If
'        Next i
'***** 2005/3/31 y.yamada del-end
        ReplaceInfo(ReplaceInfoCount).KeyString = strBuf
        '　変換後
        strKey = "変換後"
        strBuf = ReadCsvSheet(shtReplace, strKey, lngRow)
        If strBuf = "" Then
            GoTo IniErrHandler
        End If
'***** 2005/3/31 y.yamada del-str
'        For i = 0 To ReplaceInfoCount - 1
'            If strBuf = ReplaceInfo(i).KeyString And ReplaceInfo(ReplaceInfoCount).KeyString = ReplaceInfo(i).ReplaceString Then
'                GoTo IniErrHandler
'            End If
'        Next i
'***** 2005/3/31 y.yamada del-end
        ReplaceInfo(ReplaceInfoCount).ReplaceString = strBuf
        '　完全一致
        strKey = "完全一致"
        strBuf = ReadCsvSheet(shtReplace, strKey, lngRow)
        Select Case UCase(strBuf)
        Case "完全一致"
            ReplaceInfo(ReplaceInfoCount).ReplaceMode = enumReplaceMode.Complete
        Case "文字列一致"
            ReplaceInfo(ReplaceInfoCount).ReplaceMode = enumReplaceMode.Partial
        Case Else
            GoTo IniErrHandler
        End Select
        
        '整合性チェック
        For i = 0 To ReplaceInfoCount - 1
            '今回の設定値と全ての設定値を比較する
            If ReplaceInfo(i).ReplaceMode = enumReplaceMode.Partial Or ReplaceInfo(ReplaceInfoCount).ReplaceMode = enumReplaceMode.Partial Then
                '[変換前]重複チェック（一方が他方の一部になるようなパターン）
                '※[完全一致]が"完全一致"同士の場合は除く
                If InStr(ReplaceInfo(i).KeyString, ReplaceInfo(ReplaceInfoCount).KeyString) <> 0 Or InStr(ReplaceInfo(ReplaceInfoCount).KeyString, ReplaceInfo(i).KeyString) <> 0 Then
                    Call ShowIniSheet(shtReplace)
                    Call OutputMsg(MSG_003, MODE_ALL, shtReplace.Name & "#" & "変換前" & "#" & CStr(lngRow), vbCritical, APP_TITLE)
                    GoTo EndHandler
                End If
            End If
            '循環参照チェック（[変換前]と[変換後]で循環しているパターン）
            If ReplaceInfo(ReplaceInfoCount).ReplaceString = ReplaceInfo(i).KeyString And ReplaceInfo(ReplaceInfoCount).KeyString = ReplaceInfo(i).ReplaceString Then
                Call ShowIniSheet(shtReplace)
                Call OutputMsg(MSG_004, MODE_ALL, shtReplace.Name & "#" & "変換後" & "#" & CStr(lngRow), vbCritical, APP_TITLE)
                GoTo EndHandler
            End If
        Next i
        
        ReplaceInfoCount = ReplaceInfoCount + 1
    Next lngRow
    
    ReadReplaceSheet = True
EndHandler:
    On Error Resume Next
    Exit Function
IniErrHandler:      '設定エラー
    'シートをアクティブにする
    Call ShowIniSheet(shtReplace)
    Call OutputMsg(MSG_002, MODE_ALL, shtReplace.Name & "#" & strKey & "#" & CStr(lngRow), vbCritical, APP_TITLE)
    GoTo EndHandler
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "ReadReplaceSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'指定されたINIシートをアクティブにする
Private Function ShowIniSheet(sheet As Worksheet) As Boolean
    Windows(ThisWorkbook.Name).Visible = True
    Windows(ThisWorkbook.Name).WindowState = xlMaximized
    sheet.Activate
End Function

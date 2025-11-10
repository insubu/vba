Option Explicit

Private Const mModuleName As String = "basMain"

'処理対象
Private bokTarget As Workbook
Private shtTarget As Worksheet

'ヘッダー情報
Private HeaderInfo() As String
Private HeaderInfoCount As Long

'特殊列の列位置
Private PosDrepId As Long
Private PosResult As Long

'処理対象の情報
Private ServiceId As String
Private CabinetId As String
Private DrepId As String

'非致命的エラーフラグ
Private NotFatalErrorFlag As Boolean

'メニュー[アクティブシートを対象に起動]の処理
Public Function CheckSheetMain() As Boolean
    'チェック処理メイン（ファイルオープンなし）
    CheckSheetMain = CheckMain(False)
End Function

'メニュー[ファイルを指定して起動]の処理
Public Function CheckFileMain() As Boolean
    'チェック処理メイン（ファイルオープンあり）
    CheckFileMain = CheckMain(True)
End Function

'メニュー[環境設定を表示]の処理
Public Function ShowIniSheetMain() As Boolean
    On Error GoTo ErrHandler
    
    'ブックを表示する
    Windows(ThisWorkbook.Name).Visible = True
    'ブックを最大化する
    Windows(ThisWorkbook.Name).WindowState = xlMaximized
    shtMain.Activate

    ShowIniSheetMain = True
EndHandler:
    On Error Resume Next
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "ShowIniSheetMain" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'メニュー[アンインストール]の処理
Public Function UnInstallMain() As Boolean
    On Error GoTo ErrHandler
    Dim strScriptPath As String
    Dim intFileNum As Integer
    Dim lngRet As Long

    '実行確認メッセージ
    If OutputMsg(MSG_302, MODE_DLG, "", vbQuestion + vbOKCancel, APP_TITLE) = vbCancel Then
        GoTo EndHandler
    End If
    
    Application.Cursor = xlWait
    Application.StatusBar = APP_TITLE & "を削除しています..."
    Application.DisplayAlerts = False
    
    ThisWorkbook.Saved = True
    
    'メニューを削除する
    Application.CommandBars(TOOL_MENU).Delete

    'VB-Sctiptで削除スクリプトを作成する
    strScriptPath = Environ("TEMP") & "\" & objFso.GetBaseName(ThisWorkbook.FullName) & "_del.vbs"
    intFileNum = FreeFile()
    Open strScriptPath For Output As #intFileNum
    Print #intFileNum, "On Error Resume Next"
    Print #intFileNum, "Set fs = CreateObject(""Scripting.FileSystemObject"")"
    Print #intFileNum, "fs.DeleteFile """ & ThisWorkbook.FullName & """, True"
    Print #intFileNum, "fs.DeleteFile """ & strScriptPath & """, True"
    Close #intFileNum
    '削除スクリプトを実行
    lngRet = ExecuteShell(strScriptPath, vbNormalFocus)
    
    If 0 <= lngRet And lngRet <= 31 Then
        'ScriptingHostがなくてエラーになった場合バッチファイルで削除スクリプトを作成する
        strScriptPath = Environ("TEMP") & "\" & objFso.GetBaseName(ThisWorkbook.FullName) & "_del.bat"
        intFileNum = FreeFile()
        Open strScriptPath For Output As #intFileNum
        Print #intFileNum, "@echo off"
        Print #intFileNum, "del """ & ThisWorkbook.FullName & """"
        Print #intFileNum, "del """ & strScriptPath & """"
        Close #intFileNum
        '削除スクリプトを実行
        Shell strScriptPath, vbNormalFocus
    End If
    
    UnInstallMain = True
EndHandler:
    On Error Resume Next
    Close intFileNum
    Application.DisplayAlerts = True
    Application.StatusBar = False
    Application.Cursor = xlDefault
    'ブックを閉じる
    ThisWorkbook.Close
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "UnInstallMain" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'メニュー[バージョン情報]の処理
Public Function ShowVersionMain() As Boolean
    On Error GoTo ErrHandler
    
    frmVersion.AppName = APP_TITLE
    frmVersion.Major = CInt(GetCustomDocumentProperties(ThisWorkbook, "Major"))
    frmVersion.Minor = CInt(GetCustomDocumentProperties(ThisWorkbook, "Minor"))
    frmVersion.Revision = CInt(GetCustomDocumentProperties(ThisWorkbook, "Revision"))
    frmVersion.Show

    ShowVersionMain = True
EndHandler:
    On Error Resume Next
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "ShowVersionMain" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

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

'値を編集する
Private Function EditValue(strValue As String, lngAttributeInfoIndex As Long, strErrMsg As String, strEditValue As String) As Boolean
    On Error GoTo ErrHandler
    Dim strBuf As String
    Dim strChar As String
    Dim strChar_next As String
    Dim strWork As String
    Dim intPos As Integer
    Dim i As Long
'***** 2005/3/31 y.yamada ins-str
    Dim strDateStr As String
'***** 2005/3/31 y.yamada ins-end

    strBuf = strValue
    strErrMsg = ""
    strEditValue = ""
    
    '改行削除
    If AttributeInfo(lngAttributeInfoIndex).TrimCrLf = True Then
        strBuf = Replace(strBuf, vbCr, "")
        strBuf = Replace(strBuf, vbLf, "")
    End If
    
    'スペース削除
    Select Case AttributeInfo(lngAttributeInfoIndex).TrimSpace
    Case enumTrimSpaceMode.TrimAll      '■全て
        strBuf = Replace(strBuf, " ", "")
    Case enumTrimSpaceMode.TrimBoth     '■両端
'***** 2005/4/11 y.yamada upd-str
'        strBuf = Trim(strBuf)
        strBuf = TrimHankaku(strBuf)
'***** 2005/4/11 y.yamada upd-end
    Case enumTrimSpaceMode.TrimLeft     '■前方
'***** 2005/4/11 y.yamada upd-str
'        strBuf = LTrim(strBuf)
        strBuf = LTrimHankaku(strBuf)
'***** 2005/4/11 y.yamada upd-end
    Case enumTrimSpaceMode.TrimRight    '■後方
'***** 2005/4/11 y.yamada upd-str
'        strBuf = RTrim(strBuf)
        strBuf = RTrimHankaku(strBuf)
'***** 2005/4/11 y.yamada upd-end
    End Select
    
    '必須チェック
    If AttributeInfo(lngAttributeInfoIndex).Indispensable = True And strBuf = "" Then
        '必須属性未入力エラー
        strErrMsg = "必須属性が未入力です。"
        GoTo EndHandler
    End If

    '文字コードチェック
    For intPos = 1 To Len(strBuf)
        strChar = Mid(strBuf, intPos, 1)
        If IsPermittedCode(strChar) = False Then
            '文字コードエラー
            strErrMsg = "使用不可文字[" & strChar & "]が入力されています。"
            GoTo EndHandler
        End If
    Next intPos
    
    '属性の型による文字変換
    Select Case AttributeInfo(lngAttributeInfoIndex).AttrType
    Case enumAttributeType.IntegerNumber    '■整数
        '半角に変換する
        strBuf = StrConv(strBuf, vbNarrow)
'***** 2005/3/31 y.yamada ins-str
        If strBuf <> "" Then
'***** 2005/3/31 y.yamada ins-end
'***** 2005/3/31 y.yamada upd-str
'            If IsNumeric(strBuf) = False Then
            If IsNumeric(strBuf) = False Or InStr(strBuf, ",") <> 0 Then
'***** 2005/3/31 y.yamada upd-end
                strErrMsg = "数値以外が入力されています。"
                GoTo EndHandler
            End If
            If InStr(strBuf, ".") <> 0 Then
                strErrMsg = "小数点が入力されています。"
                GoTo EndHandler
            End If
'***** 2005/3/31 y.yamada ins-str
        End If
'***** 2005/3/31 y.yamada ins-end
    Case enumAttributeType.SmallNumber      '■小数
        '半角に変換する
        strBuf = StrConv(strBuf, vbNarrow)
'***** 2005/3/31 y.yamada ins-str
        If strBuf <> "" Then
'***** 2005/3/31 y.yamada ins-end
'***** 2005/3/31 y.yamada upd-str
'            If IsNumeric(strBuf) = False Then
            If IsNumeric(strBuf) = False Or InStr(strBuf, ",") <> 0 Then
'***** 2005/3/31 y.yamada upd-end
                strErrMsg = "数値以外が入力されています。"
                GoTo EndHandler
            End If
'***** 2005/4/11 y.yamada del-str
''***** 2005/3/31 y.yamada upd-str
''            If InStr(strBuf, ".") <> 0 And AttributeInfo(lngAttributeInfoIndex).ByteEditMode = enumByteEditMode.Fixed Then
'            If InStr(strBuf, ".") = 0 And AttributeInfo(lngAttributeInfoIndex).ByteEditMode <> enumByteEditMode.Complete Then
''***** 2005/3/31 y.yamada upd-end
'                strErrMsg = "小数点が入力されていません。"
'                GoTo EndHandler
'            End If
'***** 2005/3/31 y.yamada ins-str
'***** 2005/4/11 y.yamada del-end
        End If
'***** 2005/3/31 y.yamada ins-end
    Case enumAttributeType.Narrow           '■半角
        '半角に変換する
        strBuf = StrConv(strBuf, vbNarrow)
        '１文字づつチェックする
        For intPos = 1 To Len(strBuf)
            strChar = Mid(strBuf, intPos, 1)
            If IsNarrow(strChar) = False Then
                strErrMsg = "半角対象文字以外が入力されています。"
                GoTo EndHandler
            End If
        Next intPos
'***** 2005/3/31 y.yamada ins-str
    Case enumAttributeType.Date             '■日付
        '日付型の属性値を取得する
        If GetDateStr(strBuf, AttributeInfo(lngAttributeInfoIndex).DateFormat_In, strDateStr) = False Then
            strErrMsg = "入力された日付の書式が不正です。"
            GoTo EndHandler
        End If
'***** 2005/3/31 y.yamada ins-end
    Case Else
        '置換対象[完全一致]を探す
        For i = 0 To ReplaceInfoCount - 1
            If ReplaceInfo(i).ReplaceMode = enumReplaceMode.Complete Then
                If strBuf = ReplaceInfo(i).KeyString Then
                    '存在した場合は置換する
                    strBuf = ReplaceInfo(i).ReplaceString
                    Exit For
                End If
            End If
        Next i
        If i = ReplaceInfoCount Then
            '置換[完全一致]しなかった場合は１文字づづチェックする
            For intPos = 1 To Len(strBuf)
                strChar = Mid(strBuf, intPos, 1)
                '置換対象[部分一致]を探す
                For i = 0 To ReplaceInfoCount - 1
                    If ReplaceInfo(i).ReplaceMode = enumReplaceMode.Partial Then
                        If InStr(intPos, strBuf, ReplaceInfo(i).KeyString) = intPos Then
                            '存在した場合は作業用バッファに置換文字を足す
                            strWork = strWork & ReplaceInfo(i).ReplaceString
                            Exit For
                        End If
                    End If
                Next i
                If i = ReplaceInfoCount Then
                    '置換[部分一致]しなかった場合は型ごとに文字を変換する
                    Select Case AttributeInfo(lngAttributeInfoIndex).AttrType
                    Case enumAttributeType.Wide         '■全角
                        '全て全角に変換する
                        strChar = ConvToWide(strBuf, intPos)
                    Case enumAttributeType.Alphanumeric '■英数字
                        If IsAlphanumeric(strChar) = True Then
                            '英数字を半角に変換する
                            strChar = StrConv(strChar, vbNarrow)
                        Else
                            'それ以外は全角に変換する
                            strChar = ConvToWide(strBuf, intPos)
                        End If
                    Case enumAttributeType.NarrowKana   '■半角カナ
                        '半角カナを全角に変換する
                        strChar = ConvNarrowKanaToWide(strBuf, intPos)
                    End Select
                    '作業用バッファに変換後文字を足す
                    strWork = strWork & strChar
                End If
            Next intPos
            strBuf = strWork
        End If
    End Select

    'バイト数加工
    Select Case AttributeInfo(lngAttributeInfoIndex).ByteEditMode
    Case enumByteEditMode.Fixed     '■固定
        If IsCompleteByte(strBuf, AttributeInfo(lngAttributeInfoIndex).ByteSize_Left, AttributeInfo(lngAttributeInfoIndex).ByteSize_Right) = False Then
            strErrMsg = "入力された文字のバイト数が規定値と異なります。"
            GoTo EndHandler
        End If
    Case enumByteEditMode.Complete  '■補完
        If IsPermittedByte(strBuf, AttributeInfo(lngAttributeInfoIndex).ByteSize_Left, AttributeInfo(lngAttributeInfoIndex).ByteSize_Right) = False Then
            strErrMsg = "入力された文字のバイト数が規定値を超えています。"
            GoTo EndHandler
        End If
        If AttributeInfo(lngAttributeInfoIndex).AttrType = enumAttributeType.Date Then
'***** 2005/3/31 y.yamada upd-str
'            '日付型の場合は文字列の書式を変換する
'            If FormatDate(strBuf, AttributeInfo(lngAttributeInfoIndex).DateFormat_In, AttributeInfo(lngAttributeInfoIndex).DateFormat_Out, strBuf) = False Then
'                strErrMsg = "入力された日付の書式が不正です。"
'                GoTo EndHandler
'            End If
            '日付文字列の書式を変換する
            strBuf = FormatDate(strDateStr, AttributeInfo(lngAttributeInfoIndex).DateFormat_Out)
'***** 2005/3/31 y.yamada upd-end
        Else
            'それ以外の場合は文字列を指定バイト数になるように指定文字で埋める
            strBuf = FillString(strBuf, AttributeInfo(lngAttributeInfoIndex).ByteSize_Left, AttributeInfo(lngAttributeInfoIndex).ByteSize_Right, AttributeInfo(lngAttributeInfoIndex).CompleteChar)
        End If
    Case enumByteEditMode.Max       '■最大
        If IsPermittedByte(strBuf, AttributeInfo(lngAttributeInfoIndex).ByteSize_Left, AttributeInfo(lngAttributeInfoIndex).ByteSize_Right) = False Then
            strErrMsg = "入力された文字のバイト数が規定値を超えています。"
            GoTo EndHandler
        End If
    End Select
    
    '大文字/小文字統一
    Select Case AttributeInfo(lngAttributeInfoIndex).LetterType
    Case enumLetterType.Capital '■大文字
        '大文字に変換する
        strBuf = UCase(strBuf)
    Case enumLetterType.Small   '■小文字
        '小文字に変換する
        strBuf = LCase(strBuf)
    End Select

    '編集後の文字列を返す
    strEditValue = strBuf
    
    EditValue = True
EndHandler:
    On Error Resume Next
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "EditValue" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function

'***** 2005/3/31 y.yamada del-str
''日付文字列の書式を変換する
'Private Function FormatDate(strValue As String, strFormat_In As String, strFormat_Out As String, strEditValue) As Boolean
'    Dim intPos_format As Integer
'    Dim intPos_value As Integer
'    Dim strChar As String
'    Dim strChar_next As String
'    Dim strRepStr As String
'    Dim strBuf As String
'    Dim strDateStr As String
'    Dim dtmValue As Date
'
'    '半角に変換する
'    strValue = StrConv(strValue, vbNarrow)
'
'    'Date型データを作成するための雛型
'    strDateStr = "%Y/%m/%d %H:%M:%S"
'
'    intPos_value = 1
'    '入力書式を１文字づつチェックする
'    For intPos_format = 1 To Len(strFormat_In)
'        strChar = Mid(strFormat_In, intPos_format, 1)
'        If strChar = "%" And intPos_format + 1 <= Len(strFormat_In) Then
'            '"%?"を見つけた場合
'            intPos_format = intPos_format + 1
'            strRepStr = "%" & Mid(strFormat_In, intPos_format, 1)
'            If intPos_format + 1 <= Len(strFormat_In) Then
'                '次の文字を取得する
'                strChar_next = Mid(strFormat_In, intPos_format + 1, 1)
'                If InStr(intPos_value, strValue, strChar_next) = 0 Then
'                    '入力データに同じ文字が存在しない場合はエラー
'                    GoTo EndHandler
'                End If
'                '入力データから日付部分を切り出す
'                strBuf = Mid(strValue, intPos_value, InStr(intPos_value, strValue, strChar_next) - intPos_value)
'            Else
'                '入力データから日付部分を切り出す
'                strBuf = Mid(strValue, intPos_value)
'            End If
'
'            If IsNumeric(strBuf) = False Then
'                '日付部分が数値でない場合はエラー
'                GoTo EndHandler
'            End If
'
'            '取得した日付を雛型に埋め込む
'            Select Case strRepStr
'            Case "%Y"
'                If Len(strBuf) <> 4 Then
'                    '４文字以外はエラー
'                    GoTo EndHandler
'                End If
'                strDateStr = Replace(strDateStr, strRepStr, strBuf)
'            Case "%y"
'                If Len(strBuf) <> 2 Then
'                    '２文字以外はエラー
'                    GoTo EndHandler
'                End If
'                strDateStr = Replace(strDateStr, UCase(strRepStr), strBuf)
'            Case "%m", "%d", "%H", "%M", "%S"
'                If Len(strBuf) <> 1 And Len(strBuf) <> 2 Then
'                    '１or２文字以外はエラー
'                    GoTo EndHandler
'                End If
'                strDateStr = Replace(strDateStr, strRepStr, strBuf)
'            End Select
'
'            '入力データ用位置カウンタを処理した文字数分進める
'            intPos_value = intPos_value + Len(strBuf)
'        Else
'            '通常の文字の場合
'            If strChar <> Mid(strValue, intPos_value, 1) Then
'                '入力データの同じ位置に同じ文字が存在しない場合はエラー
'                GoTo EndHandler
'            End If
'
'            '入力データ用位置カウンタを処理した文字数分進める
'            intPos_value = intPos_value + 1
'        End If
'    Next intPos_format
'
'    If InStr(strDateStr, "%") <> 0 Then
'        '雛型に置換予約語"%"が残っている場合はエラー
'        GoTo EndHandler
'    End If
'
'    '雛型から日付データを作成する
'    dtmValue = CDate(strDateStr)
'
'    '出力書式に従い出力用日付文字列を作成する
'    strEditValue = strFormat_Out
'    strEditValue = Replace(strEditValue, "%Y", Format(dtmValue, "YYYY"))
'    strEditValue = Replace(strEditValue, "%y", Format(dtmValue, "YY"))
'    strEditValue = Replace(strEditValue, "%m", Format(dtmValue, "MM"))
'    strEditValue = Replace(strEditValue, "%d", Format(dtmValue, "DD"))
'    strEditValue = Replace(strEditValue, "%H", Format(dtmValue, "hh"))
'    strEditValue = Replace(strEditValue, "%M", Format(dtmValue, "mm"))
'    strEditValue = Replace(strEditValue, "%S", Format(dtmValue, "ss"))
'
'    FormatDate = True
'EndHandler:
'End Function
'***** 2005/3/31 y.yamada del-end

'***** 2005/3/31 y.yamada ins-str
'日付型の属性値を取得する
Private Function GetDateStr(ByVal strValue As String, strFormat_In As String, strDateStr As String) As Boolean
    Dim intPos_format As Integer
    Dim intPos_value As Integer
    Dim strChar As String
    Dim strChar_next As String
    Dim strRepStr As String
    Dim strBuf As String
    
    strDateStr = ""
    
    If strValue = "" Then
        '値がNULLの場合は正常終了とする
        GetDateStr = True
        GoTo EndHandler
    End If
    
    '半角に変換する
    strValue = StrConv(strValue, vbNarrow)
    
    'Date型データを作成するための雛型
    strDateStr = "%Y/%m/%d %H:%M:%S"
    
    intPos_value = 1
    '入力書式を１文字づつチェックする
    For intPos_format = 1 To Len(strFormat_In)
        strChar = Mid(strFormat_In, intPos_format, 1)
        If strChar = "%" And intPos_format + 1 <= Len(strFormat_In) Then
            '"%?"を見つけた場合
            intPos_format = intPos_format + 1
            strRepStr = "%" & Mid(strFormat_In, intPos_format, 1)
            If intPos_format + 1 <= Len(strFormat_In) Then
                '次の文字を取得する
                strChar_next = Mid(strFormat_In, intPos_format + 1, 1)
                If InStr(intPos_value, strValue, strChar_next) = 0 Then
                    '入力データに同じ文字が存在しない場合はエラー
                    GoTo EndHandler
                End If
                '入力データから日付部分を切り出す
                strBuf = Mid(strValue, intPos_value, InStr(intPos_value, strValue, strChar_next) - intPos_value)
            Else
                '入力データから日付部分を切り出す
                strBuf = Mid(strValue, intPos_value)
            End If
            
            If IsNumeric(strBuf) = False Then
                '日付部分が数値でない場合はエラー
                GoTo EndHandler
            End If
            
            '取得した日付を雛型に埋め込む
            Select Case strRepStr
            Case "%Y"
                If Len(strBuf) <> 4 Then
                    '４文字以外はエラー
                    GoTo EndHandler
                End If
                strDateStr = Replace(strDateStr, strRepStr, strBuf)
            Case "%y"
                If Len(strBuf) <> 2 Then
                    '２文字以外はエラー
                    GoTo EndHandler
                End If
                strDateStr = Replace(strDateStr, UCase(strRepStr), strBuf)
            Case "%m", "%d", "%H", "%M", "%S"
                If Len(strBuf) <> 1 And Len(strBuf) <> 2 Then
                    '１or２文字以外はエラー
                    GoTo EndHandler
                End If
                strDateStr = Replace(strDateStr, strRepStr, strBuf)
            End Select
            
            '入力データ用位置カウンタを処理した文字数分進める
            intPos_value = intPos_value + Len(strBuf)
        Else
            '通常の文字の場合
            If strChar <> Mid(strValue, intPos_value, 1) Then
                '入力データの同じ位置に同じ文字が存在しない場合はエラー
                GoTo EndHandler
            End If
            
            '入力データ用位置カウンタを処理した文字数分進める
            intPos_value = intPos_value + 1
        End If
    Next intPos_format
    
'***** 2005/4/11 y.yamada upd-str
'    If InStr(strDateStr, "%") <> 0 Then
'        '雛型に置換予約語"%"が残っている場合はエラー
'        GoTo EndHandler
'    End If
    '雛型に置換予約語"%"が残っている場合は初期値をセットする
    strDateStr = Replace(strDateStr, "%Y", Format(Now, "YYYY"))
    strDateStr = Replace(strDateStr, "%m", "1")
    strDateStr = Replace(strDateStr, "%d", "1")
    strDateStr = Replace(strDateStr, "%H", "00")
    strDateStr = Replace(strDateStr, "%M", "00")
    strDateStr = Replace(strDateStr, "%S", "00")
'***** 2005/4/11 y.yamada upd-end
    
'***** 2005/4/11 y.yamada ins-str
    If IsDate(strDateStr) = False Then
        '日付文字列として不正な場合はエラー
        GoTo EndHandler
    End If
'***** 2005/4/11 y.yamada ins-end
    
    GetDateStr = True
EndHandler:
End Function

'日付文字列の書式を変換する
Private Function FormatDate(strValue As String, strFormat_Out As String) As String
    Dim dtmValue As Date
    Dim strBuf As String

    If strValue = "" Then
        '値がNULLの場合は何もしない
        GoTo EndHandler
    End If

    '雛型から日付データを作成する
    dtmValue = CDate(strValue)

    '出力書式に従い出力用日付文字列を作成する
    strBuf = strFormat_Out
    strBuf = Replace(strBuf, "%Y", Format(dtmValue, "YYYY"))
    strBuf = Replace(strBuf, "%y", Format(dtmValue, "YY"))
    strBuf = Replace(strBuf, "%m", Format(dtmValue, "MM"))
    strBuf = Replace(strBuf, "%d", Format(dtmValue, "DD"))
    strBuf = Replace(strBuf, "%H", Format(dtmValue, "hh"))
    strBuf = Replace(strBuf, "%M", Format(dtmValue, "nn"))
    strBuf = Replace(strBuf, "%S", Format(dtmValue, "ss"))

    FormatDate = strBuf
EndHandler:
End Function
'***** 2005/3/31 y.yamada ins-end

'文字列を指定バイト数になるように指定文字で埋める
Private Function FillString(strValue As String, intByte_Left As Integer, intByte_Right As Integer, strFillStr As String) As String
    Dim strRet As String
    Dim strValue_Left As String
    Dim strValue_Right As String
    
'***** 2005/3/31 y.yamada upd-str
'    If InStr(strValue, ".") = 0 Then
'        '小数点がない場合
'        strValue_Left = strValue
'    Else
'        '小数点がある場合は整数部と小数部に分ける
'        strValue_Left = Left(strValue, InStr(strValue, ".") - 1)
'        strValue_Right = Mid(strValue, InStr(strValue, ".") + 1)
'    End If
    If intByte_Right <> 0 Then
        '小数の場合
        If InStr(strValue, ".") = 0 Then
            '小数点がない場合
            strValue_Left = strValue
        Else
            '小数点がある場合は整数部と小数部に分ける
            strValue_Left = Left(strValue, InStr(strValue, ".") - 1)
            strValue_Right = Mid(strValue, InStr(strValue, ".") + 1)
        End If
    Else
        'それ以外の場合
        strValue_Left = strValue
    End If
'***** 2005/3/31 y.yamada upd-end
    
    '整数部の文字埋めを行う
'***** 2005/3/31 y.yamada upd-str
'    If strValue_Left <> "" And intByte_Left <> 0 Then
    If intByte_Left <> 0 Then
'***** 2005/3/31 y.yamada upd-end
        strRet = FillStr(strValue_Left, strFillStr, intByte_Left, False)
    End If
    '小数部の文字埋めを行う
'***** 2005/3/31 y.yamada upd-str
'    If strValue_Right <> "" And intByte_Right <> 0 Then
    If intByte_Right <> 0 Then
'***** 2005/3/31 y.yamada upd-end
        strRet = strRet & "." & FillStr(strValue_Right, strFillStr, intByte_Right, True)
    End If
    If strRet = "" Then
        '何もしなかった場合は元の文字列を返す
        strRet = strValue
    End If

    FillString = strRet
EndHandler:
    Exit Function
End Function

'文字列のバイト数がが指定バイト数以下かどうかを取得する
Private Function IsPermittedByte(strValue As String, intByte_Left As Integer, intByte_Right As Integer) As Boolean
    Dim strValue_Left As String
    Dim strValue_Right As String
    
'***** 2005/3/31 y.yamada upd-str
'    If InStr(strValue, ".") = 0 Then
'        '小数点がない場合
'        strValue_Left = strValue
'    Else
'        '小数点がある場合は整数部と小数部に分ける
'        strValue_Left = Left(strValue, InStr(strValue, ".") - 1)
'        strValue_Right = Mid(strValue, InStr(strValue, ".") + 1)
'    End If
    If intByte_Right <> 0 Then
        '小数の場合
        If InStr(strValue, ".") = 0 Then
            '小数点がない場合
            strValue_Left = strValue
        Else
            '小数点がある場合は整数部と小数部に分ける
            strValue_Left = Left(strValue, InStr(strValue, ".") - 1)
            strValue_Right = Mid(strValue, InStr(strValue, ".") + 1)
        End If
    Else
        'それ以外の場合
        strValue_Left = strValue
    End If
'***** 2005/3/31 y.yamada upd-end
    
    '整数部をチェックする
'***** 2005/3/31 y.yamada upd-str
'    If strValue_Left <> "" And intByte_Left <> 0 Then
    If intByte_Left <> 0 Then
'***** 2005/3/31 y.yamada upd-end
        If LenB(StrConv(strValue_Left, vbFromUnicode)) > intByte_Left Then
            'バイト数オーバーエラー
            GoTo EndHandler
        End If
    End If
    '小数部をチェックする
'***** 2005/3/31 y.yamada upd-str
'    If strValue_Right <> "" And intByte_Right <> 0 Then
    If intByte_Right <> 0 Then
'***** 2005/3/31 y.yamada upd-end
        If LenB(StrConv(strValue_Right, vbFromUnicode)) > intByte_Right Then
            'バイト数オーバーエラー
            GoTo EndHandler
        End If
    End If

    IsPermittedByte = True
EndHandler:
    Exit Function
End Function

'文字列のバイト数が指定バイト数と一致するかどうかを取得する
Private Function IsCompleteByte(strValue As String, intByte_Left As Integer, intByte_Right As Integer) As Boolean
    Dim strValue_Left As String
    Dim strValue_Right As String
    
'***** 2005/3/31 y.yamada upd-str
'    If InStr(strValue, ".") = 0 Then
'        '小数点がない場合
'        strValue_Left = strValue
'    Else
'        '小数点がある場合は整数部と小数部に分ける
'        strValue_Left = Left(strValue, InStr(strValue, ".") - 1)
'        strValue_Right = Mid(strValue, InStr(strValue, ".") + 1)
'    End If
    If intByte_Right <> 0 Then
        '小数の場合
        If InStr(strValue, ".") = 0 Then
            '小数点がない場合
            strValue_Left = strValue
        Else
            '小数点がある場合は整数部と小数部に分ける
            strValue_Left = Left(strValue, InStr(strValue, ".") - 1)
            strValue_Right = Mid(strValue, InStr(strValue, ".") + 1)
        End If
    Else
        'それ以外の場合
        strValue_Left = strValue
    End If
'***** 2005/3/31 y.yamada upd-end
    
    '整数部をチェックする
'***** 2005/3/31 y.yamada upd-str
'    If strValue_Left <> "" And intByte_Left <> 0 Then
    If intByte_Left <> 0 Then
'***** 2005/3/31 y.yamada upd-end
        If LenB(StrConv(strValue_Left, vbFromUnicode)) <> intByte_Left Then
            'バイト数不一致エラー
            GoTo EndHandler
        End If
    End If
    '小数部をチェックする
'***** 2005/3/31 y.yamada upd-str
'    If strValue_Right <> "" And intByte_Right <> 0 Then
    If intByte_Right <> 0 Then
'***** 2005/3/31 y.yamada upd-end
        If LenB(StrConv(strValue_Right, vbFromUnicode)) <> intByte_Right Then
            'バイト数不一致エラー
            GoTo EndHandler
        End If
    End If

    IsCompleteByte = True
EndHandler:
    Exit Function
End Function

'指定された文字を全角に変換する
Private Function ConvToWide(strValue As String, intPos As Integer) As String
    Dim strRet As String
    Dim strChar As String
    Dim strChar_next As String
    
    strChar = Mid(strValue, intPos, 1)
    If 166 <= Asc(strChar) And Asc(strChar) <= 221 Then
        '半角カナの場合
        If intPos < Len(strValue) Then
            '次の文字が濁点・半濁点の場合は２文字分まとめて変換する
            strChar_next = Mid(strValue, intPos + 1, 1)
            If Asc(strChar_next) = 222 Or Asc(strChar_next) = 223 Then
                '濁点・半濁点の分だけ文字位置を進める
                intPos = intPos + 1
            Else
                strChar_next = ""
            End If
        End If
        strRet = StrConv(strChar & strChar_next, vbWide)
    Else
        strRet = StrConv(strChar, vbWide)
    End If
    
    ConvToWide = strRet
End Function

'指定された半角カナ文字を全角に変換する
Private Function ConvNarrowKanaToWide(strValue As String, intPos As Integer) As String
    Dim strRet As String
    Dim strChar As String
    Dim strChar_next As String
    
    strChar = Mid(strValue, intPos, 1)
    If 166 <= Asc(strChar) And Asc(strChar) <= 221 Then
        '半角カナの場合
        If intPos < Len(strValue) Then
            '次の文字が濁点・半濁点の場合は２文字分まとめて変換する
            strChar_next = Mid(strValue, intPos + 1, 1)
            If Asc(strChar_next) = 222 Or Asc(strChar_next) = 223 Then
                '濁点・半濁点の分だけ文字位置を進める
                intPos = intPos + 1
            Else
                strChar_next = ""
            End If
        End If
        strRet = StrConv(strChar & strChar_next, vbWide)
    Else
        '半角カナ以外はそのまま
        strRet = strChar
    End If
    
    ConvNarrowKanaToWide = strRet
End Function

'指定された文字が半角かどうかを取得する
Private Function IsNarrow(strChar As String) As Boolean
    Dim blnRet As Boolean
    
    If 32 <= Asc(strChar) And Asc(strChar) <= 126 Then
        'ASCIIコードが0～255
        blnRet = True       '半角である
    End If
    
    IsNarrow = blnRet
End Function

'指定された文字が英数字かどうかを取得する
Private Function IsAlphanumeric(strChar As String) As Boolean
    Dim blnRet As Boolean
    Dim strChar_work As String
    
    '半角に変換する
    strChar_work = StrConv(strChar, vbNarrow)
    If (48 <= Asc(strChar_work) And Asc(strChar_work) <= 57) Or _
       (65 <= Asc(strChar_work) And Asc(strChar_work) <= 90) Or _
       (97 <= Asc(strChar_work) And Asc(strChar_work) <= 122) Then
        'ASCIIコードが48～57・65～90・97～122
        blnRet = True       '英数字である
    End If
    
    IsAlphanumeric = blnRet
End Function

'指定された文字のコードが許容範囲かどうかを取得する
Private Function IsPermittedCode(strChar As String) As Boolean
    Dim blnRet As Boolean
    Dim i As Integer
    
    If 0 <= Asc(strChar) And Asc(strChar) <= 255 Then
        'ASCIIコードが0～255
        blnRet = True       '許容範囲である
    Else
        '句点コード情報数分繰り返す
        For i = 0 To MainInfo.PeriodCodeCount - 1
            If CInt("&H" & MainInfo.PeriodCode(i).From) <= Asc(strChar) And Asc(strChar) <= CInt("&H" & MainInfo.PeriodCode(i).To) Then
                'ASCIIコードが指定範囲内に収まる場合
                blnRet = True       '許容範囲である
                Exit For
            End If
        Next i
    End If
    
    IsPermittedCode = blnRet
End Function

'属性名をキーとして属性情報におけるインデックスを取得する
Private Function GetAttributeInfoIndex(AttrName As String) As Long
    Dim i As Long
    Dim lngIndex As Long
    
    lngIndex = -1
    For i = 0 To AttributeInfoCount - 1
        If AttributeInfo(i).AttrName = AttrName Then
            lngIndex = i
            Exit For
        End If
    Next i
    
    GetAttributeInfoIndex = lngIndex
End Function

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

'ブックを保存する
Public Sub SaveThisBook()
    On Error Resume Next
    
    If ThisWorkbook.Saved = False Then
        Application.StatusBar = APP_TITLE & "を保存しています..."
        ThisWorkbook.Save
        Application.StatusBar = False
    End If
    'ブックを非表示にする
    Windows(ThisWorkbook.Name).Visible = False
End Sub

'作業ファイルを削除する
Public Function DeleteWorkFile() As Boolean
    On Error GoTo ErrHandler
    Dim strWorkDirPath As String
    Dim strFileName As String

    Application.StatusBar = "作業ファイルを削除しています..."

    strWorkDirPath = ThisWorkbook.path & "\" & DIR_WORK
    If objFso.FolderExists(strWorkDirPath) = True Then
        'ファイルがなくなるまで繰り返す
        strFileName = Dir(strWorkDirPath & "\*")
        Do While strFileName <> ""
            Kill strWorkDirPath & "\" & strFileName
            strFileName = Dir
        Loop
    End If

    DeleteWorkFile = True
EndHandler:
    On Error Resume Next
    Application.StatusBar = False
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_LOG, mModuleName & "#" & "DeleteWorkFile" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    GoTo EndHandler
End Function



Private Const mModuleName As String = "basMessage"

Public Const MODE_LOG As Integer = 1
Public Const MODE_DLG As Integer = 2
Public Const MODE_ALL As Integer = 3

'メッセージ（$:vbCrLf / #:引数）
Public Const MSG_000 As String = "#"
Public Const MSG_001 As String = "#シート[#]の設定が不正です。"
Public Const MSG_002 As String = "#シート[#（#行目）]の設定が不正です。"
Public Const MSG_003 As String = "#シート[#（#行目）]の設定が重複しています。"
Public Const MSG_004 As String = "#シート[#（#行目）]の設定が循環参照しています。"

Public Const MSG_101 As String = "処理対象ブックが存在しません。"
Public Const MSG_102 As String = "処理対象ブックが選択されていません。"
Public Const MSG_103 As String = "対象シートは処理済みです。$再実行しますか？"
Public Const MSG_104 As String = "属性シートに属性名[#]が定義されていません。"
Public Const MSG_105 As String = "処理対象ブックが一時ファイルのためファイル保存先を特定できません。$ブックを保存するか、ファイル保存先を指定してください。"

Public Const MSG_201 As String = "エラーデータが存在します。$確認してください。"
Public Const MSG_202 As String = "チェック処理が正常に終了しました。"

Public Const MSG_301 As String = APP_TITLE & "をインストールします。"
Public Const MSG_302 As String = APP_TITLE & "をアンインストールします。"

Public Const MSG_999   As String = "予期せぬエラーが発生しました。$#-#$[#:#]"

'メッセージ情報にしたがってメッセージを出力する
'   strMsg      :メッセージ
'   intMode     :出力先(MODE_LOG:ログファイルのみ / MODE_DLG:ダイアログのみ / MODE_ALL:両方)
'   strParam    :置換文字(メッセージ内の"#"と置き換える文字列。"#"が複数ある場合はstrParamも
'                "#"で区切った文字列で渡す)
'   intStyle    :ダイアログ表示スタイル(MsgBox関数の引数"vbOkOnly"等と同様に渡す)
'   strTitle    :ダイアログのタイトル
Public Function OutputMsg(strMsg As String, intMode As Integer, Optional ByVal strParam As String = "", Optional intStyle = 0, Optional strTitle = "") As Integer
    On Error GoTo ErrHandler
    Dim strOpt()    As String       '置換文字配列
    Dim intOpt      As Integer      '置換文字配列用素数
    Dim i           As Integer      '置換文字配列のインデックス
    Dim intRet      As Integer
    
    '置換文字列を配列に格納
    If strParam <> "" Then
        Do
            ReDim Preserve strOpt(intOpt)
            If InStr(strParam, "#") = 0 Then
                strOpt(intOpt) = strParam
                intOpt = intOpt + 1
                Exit Do
            Else
                strOpt(intOpt) = Left$(strParam, InStr(strParam, "#") - 1)
                strParam = Mid$(strParam, InStr(strParam, "#") + 1)
                intOpt = intOpt + 1
            End If
        Loop
    End If
    
    '#の替わりに値を埋め込む。
    For i = 0 To intOpt - 1
        If InStr(strMsg, "#") > 0 Then
            strMsg = Replace(strMsg, "#", strOpt(i), , 1)
        End If
    Next i
    
    'ログファイルへ出力
    If intMode And MODE_LOG Then
        '$を取り除いて出力
        Call WriteLog(Replace(strMsg, "$", ""))
    End If
    
    'ダイアログへ出力
    If intMode And MODE_DLG Then
        If strTitle = "" Then strTitle = ThisWorkbook.Name
        '$を改行コードで置換して出力
        intRet = MsgBox(Replace(strMsg, "$", vbCrLf), intStyle, strTitle)
    End If
    
    OutputMsg = intRet
    
    Exit Function
ErrHandler:
    MsgBox mModuleName & "-" & "DispMsg[" & Err.number & ":" & Err.Description & "]"
End Function

'ログファイルへ出力
Private Sub WriteLog(strMsg As String)
    On Error GoTo ErrHandler
    Dim strBuf As String
    Dim intFNum As Integer
    Dim strPath As String
    
    intFNum = FreeFile()
    strPath = ThisWorkbook.path & "\" & TOOL_MENU & ".log"
    Open strPath For Append Lock Read Write As #intFNum
    
    'メッセージ書き込み
    strBuf = Format(Now, "yyyy/mm/dd hh:mm:ss")
    strBuf = strBuf & "," & strMsg
    Print #intFNum, strBuf
ErrHandler:
    Close intFNum
End Sub

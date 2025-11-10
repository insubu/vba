Public Function GetOneData(strData As String, Optional delimiter As String = ",") As String
    On Error GoTo ErrHandler
    Dim intPos As Integer       '桁位置
    Dim blnDQ As Integer        'ダブルクォーテーションの出現(0:偶数/1:奇数)
    Dim strBuf As String        '置換対象文字列
    Dim strRet As String        '切り出した文字列（１件のデータ）
    Dim blnNext As Boolean      '次のデータがあるか？（True:あり/False:なし）
    
    '区切り文字は１文字固定
    If Len(delimiter) <> 1 Then
        Err.Raise 5     '引数不正エラー
    End If
    
    'タブをトリム
    strBuf = Replace(strData, vbTab, "")
    '前スペースをトリム
    strBuf = LTrim$(strBuf)
    
    If Left$(strBuf, 1) = """" Then
        '先頭がダブルクォーテーションの場合
        For intPos = 1 To Len(strBuf)       '文字列の長さ分繰り返す
            If Mid$(strBuf, intPos, 1) = """" Then
                'ダブルクォーテーションの場合はフラグを書き換える
                blnDQ = blnDQ Xor 1
            ElseIf Mid$(strBuf, intPos, 1) = delimiter Then
                '区切り文字の場合
                If blnDQ = 0 Then
                    'ダブルクォーテーションが偶数個出現済みの場合はデータ終了と判断
                    strRet = Left$(strBuf, intPos - 1)      'データ取得
                    strData = Mid$(strBuf, intPos + 1)      '次のデータを用意
                    blnNext = True
                    Exit For
                End If
            End If
        Next intPos
    Else
        'それ以外の場合
        If InStr(strBuf, delimiter) <> 0 Then
            '区切り文字がある場合
            strRet = Left$(strBuf, InStr(strBuf, delimiter) - 1)    'データ取得
            strData = Mid$(strBuf, InStr(strBuf, delimiter) + 1)    '次のデータを用意
            blnNext = True
        End If
    End If
    
    If blnNext = False Then
        '最終項目の場合
        strRet = strData    'データ取得
        strData = ""        '次のデータを用意（空データ）
    End If
    
    '後スペースをトリム
    strRet = RTrim$(strRet)
    
    'ダブルクォーテーションをトリム（前後１個だけ）
    If Left$(strRet, 1) = """" Then strRet = Mid$(strRet, 2)
    If Right$(strRet, 1) = """" Then strRet = Left$(strRet, Len(strRet) - 1)
    '前後スペースをトリム
    strRet = Trim(strRet)

    GetOneData = strRet

    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "GetOneData" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
End Function

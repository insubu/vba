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

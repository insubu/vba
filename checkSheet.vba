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


'指定された文字が半角かどうかを取得する
Private Function IsNarrow(strChar As String) As Boolean
    Dim blnRet As Boolean
    
    If 32 <= Asc(strChar) And Asc(strChar) <= 126 Then
        'ASCIIコードが0～255
        blnRet = True       '半角である
    End If
    
    IsNarrow = blnRet
End Function

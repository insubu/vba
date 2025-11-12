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

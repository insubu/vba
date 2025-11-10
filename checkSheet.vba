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


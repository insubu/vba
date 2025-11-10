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
--------------------------------------------
def get_attribute_info_index(attr_name: str, attribute_info_list) -> int:
    """
    VBAの GetAttributeInfoIndex を Python に置き換えたもの。

    Args:
        attr_name: 検索する属性名
        attribute_info_list: AttributeInfo オブジェクトのリスト
                             各要素は obj.AttrName で参照可能

    Returns:
        int: インデックス（見つからなければ -1）
    """
    for i, attr in enumerate(attribute_info_list):
        if attr.AttrName == attr_name:
            return i
    return -1

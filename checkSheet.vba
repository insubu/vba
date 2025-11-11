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


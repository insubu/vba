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


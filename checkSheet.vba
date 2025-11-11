'置換シートを読み込む
Public Function ReadReplaceSheet() As Boolean
    On Error GoTo ErrHandler
    Dim strBuf As String
    Dim strKey As String
    Dim lngRow As Long
    Dim lngRowMax As Long
    Dim i As Long
        
    '初期化
    Erase ReplaceInfo
    ReplaceInfoCount = 0
        
'***** 2005/3/31 y.yamada upd-str
'    lngRowMax = shtReplace.Cells(Rows.Count, 1).End(xlUp).Row
    lngRowMax = shtReplace.UsedRange.Rows.Count
'***** 2005/3/31 y.yamada upd-end
    
    'シートの行数分繰り返す
    For lngRow = 3 To lngRowMax
        ReDim Preserve ReplaceInfo(ReplaceInfoCount)
        
        '設定内容をチェックする
        '　変換前
        strKey = "変換前"
        strBuf = ReadCsvSheet(shtReplace, strKey, lngRow)
        If strBuf = "" Then
'***** 2005/3/31 y.yamada upd-str
'            GoTo IniErrHandler
            '未入力の場合は有効行でないと判断する
            Exit For
'***** 2005/3/31 y.yamada upd-end
        End If
'***** 2005/3/31 y.yamada del-str
'        For i = 0 To ReplaceInfoCount - 1
'            If InStr(ReplaceInfo(i).KeyString, strBuf) <> 0 Or InStr(strBuf, ReplaceInfo(i).KeyString) <> 0 Then
'                GoTo IniErrHandler
'            End If
'        Next i
'***** 2005/3/31 y.yamada del-end
        ReplaceInfo(ReplaceInfoCount).KeyString = strBuf
        '　変換後
        strKey = "変換後"
        strBuf = ReadCsvSheet(shtReplace, strKey, lngRow)
        If strBuf = "" Then
            GoTo IniErrHandler
        End If
'***** 2005/3/31 y.yamada del-str
'        For i = 0 To ReplaceInfoCount - 1
'            If strBuf = ReplaceInfo(i).KeyString And ReplaceInfo(ReplaceInfoCount).KeyString = ReplaceInfo(i).ReplaceString Then
'                GoTo IniErrHandler
'            End If
'        Next i
'***** 2005/3/31 y.yamada del-end
        ReplaceInfo(ReplaceInfoCount).ReplaceString = strBuf
        '　完全一致
        strKey = "完全一致"
        strBuf = ReadCsvSheet(shtReplace, strKey, lngRow)
        Select Case UCase(strBuf)
        Case "完全一致"
            ReplaceInfo(ReplaceInfoCount).ReplaceMode = enumReplaceMode.Complete
        Case "文字列一致"
            ReplaceInfo(ReplaceInfoCount).ReplaceMode = enumReplaceMode.Partial
        Case Else
            GoTo IniErrHandler
        End Select
        
        '整合性チェック
        For i = 0 To ReplaceInfoCount - 1
            '今回の設定値と全ての設定値を比較する
            If ReplaceInfo(i).ReplaceMode = enumReplaceMode.Partial Or ReplaceInfo(ReplaceInfoCount).ReplaceMode = enumReplaceMode.Partial Then
                '[変換前]重複チェック（一方が他方の一部になるようなパターン）
                '※[完全一致]が"完全一致"同士の場合は除く
                If InStr(ReplaceInfo(i).KeyString, ReplaceInfo(ReplaceInfoCount).KeyString) <> 0 Or InStr(ReplaceInfo(ReplaceInfoCount).KeyString, ReplaceInfo(i).KeyString) <> 0 Then
                    Call ShowIniSheet(shtReplace)
                    Call OutputMsg(MSG_003, MODE_ALL, shtReplace.Name & "#" & "変換前" & "#" & CStr(lngRow), vbCritical, APP_TITLE)
                    GoTo EndHandler
                End If
            End If
            '循環参照チェック（[変換前]と[変換後]で循環しているパターン）
            If ReplaceInfo(ReplaceInfoCount).ReplaceString = ReplaceInfo(i).KeyString And ReplaceInfo(ReplaceInfoCount).KeyString = ReplaceInfo(i).ReplaceString Then
                Call ShowIniSheet(shtReplace)
                Call OutputMsg(MSG_004, MODE_ALL, shtReplace.Name & "#" & "変換後" & "#" & CStr(lngRow), vbCritical, APP_TITLE)
                GoTo EndHandler
            End If
        Next i
        
        ReplaceInfoCount = ReplaceInfoCount + 1
    Next lngRow
    
    ReadReplaceSheet = True
EndHandler:
    On Error Resume Next
    Exit Function
IniErrHandler:      '設定エラー
    'シートをアクティブにする
    Call ShowIniSheet(shtReplace)
    Call OutputMsg(MSG_002, MODE_ALL, shtReplace.Name & "#" & strKey & "#" & CStr(lngRow), vbCritical, APP_TITLE)
    GoTo EndHandler
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "ReadReplaceSheet" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
    Resume EndHandler
End Function


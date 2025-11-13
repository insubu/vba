                        If InStr(intPos, strBuf, ReplaceInfo(i).KeyString) = intPos Then
                            '存在した場合は作業用バッファに置換文字を足す
                            strWork = strWork & ReplaceInfo(i).ReplaceString
                            Exit For
                        End If

'日付文字列の書式を変換する
Private Function FormatDate(strValue As String, strFormat_Out As String) As String
    Dim dtmValue As Date
    Dim strBuf As String

    If strValue = "" Then
        '値がNULLの場合は何もしない
        GoTo EndHandler
    End If

    '雛型から日付データを作成する
    dtmValue = CDate(strValue)

    '出力書式に従い出力用日付文字列を作成する
    strBuf = strFormat_Out
    strBuf = Replace(strBuf, "%Y", Format(dtmValue, "YYYY"))
    strBuf = Replace(strBuf, "%y", Format(dtmValue, "YY"))
    strBuf = Replace(strBuf, "%m", Format(dtmValue, "MM"))
    strBuf = Replace(strBuf, "%d", Format(dtmValue, "DD"))
    strBuf = Replace(strBuf, "%H", Format(dtmValue, "hh"))
    strBuf = Replace(strBuf, "%M", Format(dtmValue, "nn"))
    strBuf = Replace(strBuf, "%S", Format(dtmValue, "ss"))

    FormatDate = strBuf
EndHandler:
End Function
'***** 2005/3/31 y.yamada ins-end


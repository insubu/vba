Option Explicit

'========================================
' メイン処理
'========================================
Sub ScheduleByEmployee_WorkdayOnly()

    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim startRow As Long
    startRow = 2                 ' データ開始行
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim startCol As Long
    startCol = ws.Range("M1").Column   ' カレンダー開始列
    
    Dim empDict As Object
    Set empDict = CreateObject("Scripting.Dictionary")
    
    Dim r As Long
    
    '------------------------------------
    ' 社員名 → 行番号の対応表を作成
    '------------------------------------
    For r = startRow To lastRow
        If Not empDict.Exists(ws.Cells(r, "A").Value) Then
            empDict.Add ws.Cells(r, "A").Value, New Collection
        End If
        empDict(ws.Cells(r, "A").Value).Add r
    Next r
    
    Dim emp As Variant
    
    '====================================
    ' 社員ごとに排程
    '====================================
    For Each emp In empDict.Keys
        
        Dim col As Long
        col = startCol
        
        '--------------------------------
        ' 第1フェーズ：工数 >= 1 を優先配置
        '--------------------------------
        For Each r In empDict(emp)
            
            Dim work As Double
            work = ws.Cells(r, "C").Value
            
            If work >= 1 Then
                Do While work >= 1
                    col = NextWorkCol(ws, col)
                    ws.Cells(r, col).Value = 1
                    work = work - 1
                    col = col + 1
                Loop
                
                ' 小数部分を工数列に残す
                ws.Cells(r, "C").Value = work
            End If
            
        Next r
        
        '--------------------------------
        ' 第2フェーズ：工数 < 1 を最後の稼働日に配置
        '--------------------------------
        col = NextWorkCol(ws, col)
        
        For Each r In empDict(emp)
            work = ws.Cells(r, "C").Value
            If work > 0 Then
                ws.Cells(r, col).Value = work
            End If
        Next r
        
    Next emp
    
End Sub

'========================================
' 指定列が稼働日かどうか判定
'========================================
Function IsWorkDay(ws As Worksheet, col As Long) As Boolean
    
    Dim d As Date
    d = ws.Cells(1, col).Value   ' 日付は1行目
    
    ' 土日判定（月曜始まり）
    If Weekday(d, vbMonday) >= 6 Then
        IsWorkDay = False
        Exit Function
    End If
    
    ' 背景色が緑（休息日）
    If ws.Cells(1, col).Interior.Color = RGB(0, 255, 0) Then
        IsWorkDay = False
        Exit Function
    End If
    
    IsWorkDay = True
    
End Function

'========================================
' 次の稼働日列を取得
'========================================
Function NextWorkCol(ws As Worksheet, col As Long) As Long
    
    Do While Not IsWorkDay(ws, col)
        col = col + 1
    Loop
    
    NextWorkCol = col
    
End Function

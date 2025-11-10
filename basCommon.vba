'ファイル選択ダイアログを表示
Public Function OpenFileDialog(Optional strDirPath As String = "", Optional strTitle As String = "", Optional strFilter As String = "") As String
    Dim udtOpenFileName As OPENFILENAME
    Dim lngRet As Long
    
    '引数をセット
    With udtOpenFileName
        .lStructSize = Len(udtOpenFileName)
        .lpstrInitialDir = strDirPath           '初期表示ディレクトリ
        .lpstrTitle = strTitle                  'ダイアログのタイトル
        .lpstrFilter = strFilter                'フィルタ文字列
        .nMaxFile = 256                         'ファイル名の最大長（パス含む）
        .lpstrFile = String(256, vbNullChar)    'ファイル名を格納する文字列
        'オプション（未存在ファイル名入力不可・読み取り専用チェックOFF）
        .flags = OFN_FILEMUSTEXIST Or OFN_HIDEREADONLY
    End With
    'ファイル選択ダイアログを表示
    lngRet = GetOpenFileName(udtOpenFileName)
    
    OpenFileDialog = Left(udtOpenFileName.lpstrFile, InStr(udtOpenFileName.lpstrFile, vbNullChar) - 1)

End Function

'ファイル保存ダイアログを表示
Public Function SaveFileDialog(Optional strDirPath As String = "", Optional strTitle As String = "", Optional strFilter As String = "", Optional strExtention As String = "") As String
    Dim udtOpenFileName As OPENFILENAME
    Dim lngRet As Long
    
    '引数をセット
    With udtOpenFileName
        .lStructSize = Len(udtOpenFileName)
        .lpstrInitialDir = strDirPath           '初期表示ディレクトリ
        .lpstrTitle = strTitle                  'ダイアログのタイトル
        .lpstrFilter = strFilter                'フィルタ文字列
        .lpstrDefExt = strExtention             '省略時の拡張子
        .nMaxFile = 256                         'ファイル名の最大長（パス含む）
        .lpstrFile = String(256, vbNullChar)    'ファイル名を格納する文字列
        'オプション（読み取り専用チェックOFF）
        .flags = OFN_HIDEREADONLY
    End With
    'ファイル選択ダイアログを表示
    lngRet = GetOpenFileName(udtOpenFileName)
    
    SaveFileDialog = Left(udtOpenFileName.lpstrFile, InStr(udtOpenFileName.lpstrFile, vbNullChar) - 1)

End Function

'フォルダ選択ダイアログを表示
Public Function FolderDialog(hwndOwner As Long, Optional strTitle As String = "", Optional strDirPath As String = "") As String
    Dim udtBrowseInfo         As BROWSEINFO
    Dim lngPidl               As Long
    Dim strPath               As String * MAX_PATH
    Dim lngWin32apiResultCode As Long

    '引数をセット
    With udtBrowseInfo
        .hwndOwner = hwndOwner
        .lpszTitle = strTitle
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_DONTGOBELOWDOMAIN Or _
                   BIF_STATUSTEXT Or BIF_EDITBOX Or _
                   BIF_BROWSEINCLUDEFILES Or BIF_VALIDATE
    
        .lpfn = FARPROC(AddressOf BrowseCallbackProc)       'コールバック関数のアドレス
        .lParam = strDirPath & vbNullChar                   '初期フォルダのパス名
    End With
    'フォルダの選択ダイアログボックスを表示
    lngPidl = SHBrowseForFolder(udtBrowseInfo)
    'ユーザーが選択したときは
    If lngPidl Then
        '選択結果をファイルシステムパスへ変換
        lngWin32apiResultCode = SHGetPathFromIDList(ByVal lngPidl, strPath)
        '選択結果を表示
        FolderDialog = Left(strPath, InStr(strPath, vbNullChar) - 1)
        'ITEMIDLISTを解放
        CoTaskMemFree ByVal lngPidl
    End If

End Function
Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If uMsg = BFFM_INITIALIZED Then
        SendMessage hwnd, BFFM_SETSELECTIONA, 1, ByVal lpData
    End If
End Function
Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

'ファイルを関連付けされたアプリで開く
Public Function ExecuteShell(strFilePath As String, lngStyle As Long) As Long
    Dim lngWin32apiResultCode As Long
    
    ' ファイルを関連付けされたアプリで開く
    lngWin32apiResultCode = ShellExecute(GetForegroundWindow(), "open", strFilePath, vbNullChar, vbNullChar, lngStyle)
    ExecuteShell = lngWin32apiResultCode
End Function

'文字列の両端にダブルコーテーションを付加
Public Function PutDQ(strData As String) As String
    PutDQ = """" & strData & """"
End Function

'タブ・スペースをトリム
Public Function TrimEx(strData As String) As String
    Dim intLPos As Integer
    Dim intRPos As Integer
    
    '前方検索してタブ or スペースでなくなる位置を取得
    For intLPos = 1 To Len(strData)
        If Mid(strData, intLPos, 1) <> vbTab And Mid(strData, intLPos, 1) <> " " Then Exit For
    Next intLPos
    
    '後方検索してタブ or スペースでなくなる位置を取得
    For intRPos = Len(strData) To 1 Step -1
        If Mid(strData, intRPos, 1) <> vbTab And Mid(strData, intRPos, 1) <> " " Then Exit For
    Next intRPos
  
    'トリムした文字列を返す
    If intRPos >= intLPos Then TrimEx = Mid(strData, intLPos, intRPos - intLPos + 1)

End Function

'***** 2005/4/11 y.yamada ins-str
'文字列前方の半角スペースをトリム
Public Function LTrimHankaku(strData As String) As String
    Dim intPos As Integer
    
    '前方検索して半角スペースでなくなる位置を取得
    For intPos = 1 To Len(strData)
        If Mid(strData, intPos, 1) <> " " Then Exit For
    Next intPos
  
    'トリムした文字列を返す
    LTrimHankaku = Mid(strData, intPos)
End Function

'文字列後方の半角スペースをトリム
Public Function RTrimHankaku(strData As String) As String
    Dim intPos As Integer
    
    '後方検索して半角スペースでなくなる位置を取得
    For intPos = Len(strData) To 1 Step -1
        If Mid(strData, intPos, 1) <> " " Then Exit For
    Next intPos
  
    'トリムした文字列を返す
    RTrimHankaku = Left(strData, intPos)
End Function

'文字列の半角スペースをトリム
Public Function TrimHankaku(strData As String) As String
    Dim strBuf As String
  
    strBuf = strData
    '文字列前方の半角スペースをトリム
    strBuf = LTrimHankaku(strBuf)
    '文字列後方の半角スペースをトリム
    strBuf = RTrimHankaku(strBuf)
  
    'トリムした文字列を返す
    TrimHankaku = strBuf
End Function
'***** 2005/4/11 y.yamada ins-end

'文字列をMIME形式(「x-www-form-url 符号化」形式)に変換する
Public Function UrlEncode(sSrc As String) As String
    On Error GoTo ErrHandler
    Dim strOut As String
    Dim strChar As String
    Dim strByte As String
    Dim intCharIndex As Integer
    Dim intByteIndex As Integer
    
    strOut = ""
    For intCharIndex = 1 To Len(sSrc)
        strChar = Mid(sSrc, intCharIndex, 1)
        If (("a" <= strChar) And (strChar <= "z")) _
        Or (("A" <= strChar) And (strChar <= "Z")) _
        Or (("0" <= strChar) And (strChar <= "9")) Then
            '半角英数字はそのまま
            strOut = strOut & strChar
        ElseIf strChar = " " Then
            '半角スペースは"+"に置きかえる
            strOut = strOut & "+"
        Else
            'それ以外はエンコードする
            strChar = StrConv(strChar, vbFromUnicode)
            For intByteIndex = 1 To LenB(strChar)
                strByte = MidB(strChar, intByteIndex, 1)
                strOut = strOut & "%" & Right("0" & Hex(AscB(strByte)), 2)
            Next intByteIndex
        End If
    Next intCharIndex
    
    UrlEncode = strOut
    Exit Function
ErrHandler:
    Call OutputMsg(MSG_999, MODE_ALL, mModuleName & "#" & "UrlEncode" & "#" & Err.number & "#" & Err.Description, vbCritical, APP_TITLE)
End Function
-------------------------------------------------------------

import ctypes
import os
import urllib.parse
from tkinter import filedialog, Tk


# =========================================================
# ファイル選択ダイアログを表示
# =========================================================
def open_file_dialog(dir_path: str = "", title: str = "", filetypes: str = "") -> str:
    root = Tk()
    root.withdraw()  # 隠す
    file_path = filedialog.askopenfilename(
        initialdir=dir_path or os.getcwd(),
        title=title or "ファイルを選択してください",
        filetypes=[("すべてのファイル", "*.*")] if not filetypes else parse_filter(filetypes)
    )
    root.destroy()
    return file_path or ""


# =========================================================
# ファイル保存ダイアログを表示
# =========================================================
def save_file_dialog(dir_path: str = "", title: str = "", filetypes: str = "", ext: str = "") -> str:
    root = Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(
        initialdir=dir_path or os.getcwd(),
        title=title or "ファイルを保存",
        defaultextension=ext,
        filetypes=[("すべてのファイル", "*.*")] if not filetypes else parse_filter(filetypes)
    )
    root.destroy()
    return file_path or ""


# =========================================================
# フォルダ選択ダイアログを表示
# =========================================================
def folder_dialog(title: str = "", dir_path: str = "") -> str:
    root = Tk()
    root.withdraw()
    folder = filedialog.askdirectory(
        title=title or "フォルダを選択してください",
        initialdir=dir_path or os.getcwd()
    )
    root.destroy()
    return folder or ""


# =========================================================
# ファイルを関連付けされたアプリで開く
# =========================================================
def execute_shell(file_path: str, style: int = 1) -> int:
    ShellExecute = ctypes.windll.shell32.ShellExecuteW
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    result = ShellExecute(hwnd, "open", file_path, None, None, style)
    return result


# =========================================================
# 文字列の両端にダブルクォーテーションを付加
# =========================================================
def put_dq(data: str) -> str:
    return f'"{data}"'


# =========================================================
# タブ・スペースをトリム
# =========================================================
def trim_ex(data: str) -> str:
    return data.strip(" \t")


# =========================================================
# 半角スペース専用トリム
# =========================================================
def ltrim_hankaku(data: str) -> str:
    return data.lstrip(" ")


def rtrim_hankaku(data: str) -> str:
    return data.rstrip(" ")


def trim_hankaku(data: str) -> str:
    return data.strip(" ")


# =========================================================
# MIME形式(x-www-form-urlencoded)に変換
# =========================================================
def url_encode(s: str) -> str:
    try:
        # urllib.parse.quote_plus 可直接实现同样效果
        return urllib.parse.quote_plus(s, safe="")
    except Exception as e:
        print(f"[UrlEncode Error] {e}")
        return ""


# =========================================================
# 補助：VBAのFilter文字列をtkinter用に変換
# =========================================================
def parse_filter(vba_filter: str):
    """
    VBAのFilter形式（例: "CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*"）
    を tkinter の filetypes=[("CSVファイル","*.csv"), ("すべてのファイル","*.*")]
    に変換する
    """
    parts = vba_filter.split("|")
    result = []
    for i in range(0, len(parts) - 1, 2):
        result.append((parts[i], parts[i + 1]))
    return result

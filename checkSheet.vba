from dataclasses import dataclass, field
from enum import Enum
import os


# --- Enums and data classes ---

class SaveMode(Enum):
    Csv = "CSV"
    TextTab = "TEXT(ﾀﾌﾞ)"
    TextComma = "TEXT(ｶﾝﾏ)"
    Fixed = "固定長"


class TrimSpaceMode(Enum):
    Non = 0
    TrimAll = 1
    TrimLeft = 2
    TrimRight = 3
    TrimBoth = 4


@dataclass
class OriginCellInfo:
    Row: int = 1
    Col: int = 1
    DeleteUpperRow: bool = False
    AddHeader: bool = False


@dataclass
class PeriodCode:
    FromCode: str = ""
    ToCode: str = ""


@dataclass
class MainSheetInfo:
    SaveDirPath: str = ""
    SaveMode: SaveMode = None
    SaveExtension: str = ""
    OriginCell: OriginCellInfo = field(default_factory=OriginCellInfo)
    TextSeparator: str = ","
    ErrCellColor: str = ""
    EditCellColor: str = ""
    ResultHeaderName: str = ""
    TrimHeaderSpace: TrimSpaceMode = TrimSpaceMode.Non
    TrimHeaderCrLf: bool = False
    PeriodCodes: list[PeriodCode] = field(default_factory=list)


# --- Main read function ---

def read_main_sheet(read_func, show_func, msg_func) -> MainSheetInfo | None:
    """
    Read 'main sheet' configuration using a provided read_func(key)->value function.
    :param read_func: callable that retrieves a value by key (like ReadIniSheet in VBA)
    :param show_func: callable to show a sheet (like ShowIniSheet)
    :param msg_func: callable to output a message (like OutputMsg)
    :return: MainSheetInfo or None if error
    """
    info = MainSheetInfo()
    try:
        # --- ファイル保存先 ---
        key = "ファイル保存先"
        buf = read_func(key)
        if buf and not os.path.isdir(buf):
            raise ValueError(f"Invalid save directory: {buf}")
        info.SaveDirPath = buf

        # --- ファイル保存方法 ---
        key = "ファイル保存方法"
        work = read_func(key)
        mode = get_one_data(work, ":").upper()
        if mode == "CSV":
            info.SaveMode = SaveMode.Csv
        elif mode in ("TEXT(ﾀﾌﾞ)", "TEXT(TAB)"):
            info.SaveMode = SaveMode.TextTab
        elif mode in ("TEXT(ｶﾝﾏ)", "TEXT(COMMA)"):
            info.SaveMode = SaveMode.TextComma
        elif mode == "固定長":
            info.SaveMode = SaveMode.Fixed
        else:
            raise ValueError(f"Invalid save mode: {mode}")

        # --- 保存ファイル拡張子 ---
        buf = get_one_data(work, ":")
        if not buf:
            raise ValueError("Missing save extension")
        info.SaveExtension = buf

        # --- 開始セル ---
        key = "開始セル"
        work = read_func(key)
        row = get_one_data(work, ",")
        col = get_one_data(work, ":")
        if not row.isdigit() or not col.isdigit():
            raise ValueError("Invalid origin cell")
        info.OriginCell.Row = int(row)
        info.OriginCell.Col = int(col)

        # --- 開始行より上の行を削除 ---
        del_flag = get_one_data(work, ":").upper()
        info.OriginCell.DeleteUpperRow = del_flag == "Y"

        # --- ヘッダ追加 ---
        header_flag = get_one_data(work, ":").upper()
        info.OriginCell.AddHeader = header_flag == "ADD"

        # --- セパレータ ---
        key = "セパレータ"
        sep = read_func(key)
        if not sep:
            raise ValueError("Missing separator")
        info.TextSeparator = "\t" if sep.upper() == "TAB" else sep

        # --- エラー背景色 / 編集済背景色 ---
        info.ErrCellColor = read_func("エラー背景色")
        info.EditCellColor = read_func("編集済背景色")

        # --- 処理結果 ---
        info.ResultHeaderName = read_func("処理結果")

        # --- ヘッダー行スペース削除 ---
        trim_mode = read_func("ヘッダー行スペース削除")
        trim_map = {
            "": TrimSpaceMode.Non,
            "全て": TrimSpaceMode.TrimAll,
            "前方": TrimSpaceMode.TrimLeft,
            "後方": TrimSpaceMode.TrimRight,
            "両端": TrimSpaceMode.TrimBoth,
        }
        info.TrimHeaderSpace = trim_map.get(trim_mode, TrimSpaceMode.Non)

        # --- ヘッダー行改行削除 ---
        info.TrimHeaderCrLf = read_func("ヘッダー行改行削除").upper() == "Y"

        # --- 句点コード ---
        work = read_func("句点コード")
        while work:
            code = get_one_data(work, ",")
            frm = get_one_data(code, ":")
            to = get_one_data(code, ":") or frm
            info.PeriodCodes.append(PeriodCode(frm, to))

        return info

    except Exception as e:
        show_func("shtMain")
        msg_func(f"設定エラー: {key}, {e}")
        return None


# --- Helpers ---

def get_one_data(s: str, delim: str) -> str:
    """Pop the first token from a string based on delimiter."""
    if not s:
        return ""
    parts = s.split(delim, 1)
    first = parts[0].strip()
    if len(parts) == 2:
        globals()["last_read_buffer"] = parts[1]
    else:
        globals()["last_read_buffer"] = ""
    return first

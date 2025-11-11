from dataclasses import dataclass, field
from enum import Enum


# --- ENUM DEFINITIONS (VBA equivalents) ---

class AttributeType(Enum):
    Non = 0
    Narrow = 1
    Wide = 2
    Alphanumeric = 3
    NarrowKana = 4
    IntegerNumber = 5
    SmallNumber = 6
    Date = 7


class LetterType(Enum):
    Non = 0
    Capital = 1
    Small = 2


class ByteEditMode(Enum):
    Non = 0
    Fixed = 1
    Max = 2
    Complete = 3


class TrimSpaceMode(Enum):
    Non = 0
    TrimAll = 1
    TrimLeft = 2
    TrimRight = 3
    TrimBoth = 4


# --- ATTRIBUTE INFO STRUCTURE ---

@dataclass
class AttributeInfoItem:
    AttrName: str = ""
    ColPos: int = 0
    Indispensable: bool = False
    AttrType: AttributeType = AttributeType.Non
    DateFormat_In: str = ""
    DateFormat_Out: str = ""
    LetterType: LetterType = LetterType.Non
    ByteSize_Left: int = 0
    ByteSize_Right: int = 0
    ByteEditMode: ByteEditMode = ByteEditMode.Non
    CompleteChar: str = ""
    TrimSpace: TrimSpaceMode = TrimSpaceMode.Non
    TrimCrLf: bool = False


# --- CORE FUNCTION ---

def read_attribute_sheet(
    read_func,          # like ReadCsvSheet(sheet, key, row)
    show_func,          # like ShowIniSheet(sheet)
    msg_func,           # like OutputMsg()
    sheet_name: str,
    max_rows: int,
    add_header_mode: bool
) -> list[AttributeInfoItem] | None:
    """
    Reads attribute sheet definitions from a source function.

    Args:
        read_func: callable(sheet_name, key, row) -> str
        show_func: callable(sheet_name)
        msg_func: callable(message)
        sheet_name: sheet name (for error message)
        max_rows: number of rows to process
        add_header_mode: bool (from MainInfo.OriginCell.AddHeader)

    Returns:
        list[AttributeInfoItem] or None if error
    """
    attrs: list[AttributeInfoItem] = []
    try:
        for row in range(3, max_rows + 1):
            # --- 属性名 ---
            key = "属性名"
            buf = read_func(sheet_name, key, row)
            if not buf:
                # 空白行なら終了（=有効行終了）
                break
            attr = AttributeInfoItem(AttrName=buf)

            # --- 属性位置 ---
            key = "属性位置"
            buf = read_func(sheet_name, key, row)
            if add_header_mode:
                if not buf.isdigit():
                    raise ValueError(f"Invalid column position '{buf}' at row {row}")
                attr.ColPos = int(buf)

            # --- 必須 ---
            key = "必須"
            buf = read_func(sheet_name, key, row).upper()
            if buf == "Y":
                attr.Indispensable = True
            elif buf == "N":
                attr.Indispensable = False
            else:
                raise ValueError(f"Invalid 必須 flag '{buf}'")

            # --- 型 ---
            key = "型"
            buf = read_func(sheet_name, key, row)
            attr.AttrType, attr.DateFormat_In = parse_attr_type(buf)

            # --- 大文字/小文字 ---
            key = "大文字/小文字"
            buf = read_func(sheet_name, key, row)
            attr.LetterType = {
                "": LetterType.Non,
                "大文字": LetterType.Capital,
                "小文字": LetterType.Small
            }.get(buf, None)
            if attr.LetterType is None:
                raise ValueError(f"Invalid 大文字/小文字 '{buf}'")

            # --- バイト数 ---
            key = "バイト数"
            buf = read_func(sheet_name, key, row)
            if buf:
                if not is_number_format(buf):
                    raise ValueError(f"Invalid バイト数 '{buf}'")
                if "." in buf:
                    left, right = buf.split(".", 1)
                    attr.ByteSize_Left = int(left)
                    attr.ByteSize_Right = int(right)
                else:
                    attr.ByteSize_Left = int(buf)

            # --- バイト数加工 ---
            key = "バイト数加工"
            buf = read_func(sheet_name, key, row)
            if buf:
                parse_byte_edit_mode(attr, buf)

            # --- スペース削除 ---
            key = "スペース削除"
            buf = read_func(sheet_name, key, row)
            trim_map = {
                "": TrimSpaceMode.Non,
                "全て": TrimSpaceMode.TrimAll,
                "前方": TrimSpaceMode.TrimLeft,
                "後方": TrimSpaceMode.TrimRight,
                "両端": TrimSpaceMode.TrimBoth,
            }
            if buf not in trim_map:
                raise ValueError(f"Invalid スペース削除 '{buf}'")
            attr.TrimSpace = trim_map[buf]

            # --- 改行削除 ---
            key = "改行削除"
            buf = read_func(sheet_name, key, row).upper()
            if buf == "Y":
                attr.TrimCrLf = True
            elif buf == "N":
                attr.TrimCrLf = False
            else:
                raise ValueError(f"Invalid 改行削除 '{buf}'")

            attrs.append(attr)

        return attrs

    except Exception as e:
        show_func(sheet_name)
        msg_func(f"設定エラー: {sheet_name}#{key}#{row} - {e}")
        return None


# --- HELPER FUNCTIONS ---

def parse_attr_type(buf: str):
    """Return (AttributeType, DateFormat_In)"""
    if not buf:
        return AttributeType.Non, ""

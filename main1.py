from dataclasses import dataclass
from pathlib import Path
import pandas as pd
import shutil

@dataclass
class OriginCell:
    row: int
    col: int
    add_header: bool = False
    delete_upper_row: bool = False


@dataclass
class MainInfo:
    text_separator: str = ','
    result_header_name: str = '結果'
    origin_cell: OriginCell = OriginCell(1, 1)
    trim_header_crlf: bool = True
    trim_header_space: str = 'TrimBoth'
    err_cell_color: str = '#ffcccc'
    edit_cell_color: str = '#ccffcc'
    save_dir_path: str = ''


class BasMain:
    def __init__(self, main_info: MainInfo):
        self.main_info = main_info
        self.target_file: Path | None = None
        self.df: pd.DataFrame | None = None

    def check_file_main(self, file_path: str):
        """菜单: 文件を指定して起動"""
        self.target_file = Path(file_path)
        if not self.target_file.exists():
            raise FileNotFoundError(file_path)

        # 拷贝至工作目录
        work_dir = Path(__file__).parent / "work"
        work_dir.mkdir(exist_ok=True)
        work_file = work_dir / f"{self.target_file.stem}.txt"
        shutil.copy(self.target_file, work_file)

        # 读取文本（判定分隔符）
        sep = self._detect_separator(self.main_info.text_separator)
        self.df = pd.read_csv(work_file, sep=sep, dtype=str).fillna('')
        print(f"Loaded {len(self.df)} rows from {work_file}")

        # 检查
        self.check_sheet()
        self.save_result_to_file(work_file.with_suffix('.checked.csv'))
        print("✅ Check complete")

    def _detect_separator(self, sep_text):
        if sep_text == "\\t" or sep_text.lower() == "tab":
            return "\t"
        elif sep_text in (";", ",", " "):
            return sep_text
        else:
            return sep_text  # custom

    def check_sheet(self):
        """模仿 VBA 的 CheckSheet 主体"""
        for idx, row in self.df.iterrows():
            # 这里只示范整数型检查逻辑
            for col, value in row.items():
                if col.endswith("_id") and not value.isdigit():
                    self.df.at[idx, col] = f"NG[{col}]"
        self.df["結果"] = "OK"

    def save_result_to_file(self, out_path: Path):
        """保存结果到文件"""
        self.df.to_csv(out_path, index=False, encoding='utf-8-sig')
        print(f"Saved checked file: {out_path}")

from datetime import datetime
import re

def get_date_str(value: str, format_in: str) -> tuple[bool, str]:
    """
    GetDateStr VBA版のPython移植
    指定書式(format_in)に従って value(日付文字列)を検証・正規化し、
    "%Y/%m/%d %H:%M:%S" 形式の文字列を返す。

    Returns:
        (success: bool, date_str: str)
    """
    if not value:
        return True, ""  # 空値は正常終了

    # 半角化（PythonにはvbNarrowがないため単純に全角数字→半角）
    z2h_table = str.maketrans("０１２３４５６７８９", "0123456789")
    value = value.translate(z2h_table)

    date_str = "%Y/%m/%d %H:%M:%S"
    pos = 0

    i = 0
    while i < len(format_in):
        ch = format_in[i]
        if ch == "%" and i + 1 < len(format_in):
            i += 1
            rep = "%" + format_in[i]
            next_char = format_in[i + 1] if i + 1 < len(format_in) else ""
            # 次の区切り文字までを取得
            if next_char and next_char in value[pos:]:
                next_pos = value.index(next_char, pos)
                buf = value[pos:next_pos]
            else:
                buf = value[pos:]

            if not buf.isdigit():
                return False, date_str

            if rep == "%Y":
                if len(buf) != 4:
                    return False, date_str
                date_str = date_str.replace(rep, buf)
            elif rep == "%y":
                if len(buf) != 2:
                    return False, date_str
                date_str = date_str.replace(rep.upper(), buf)
            elif rep in ("%m", "%d", "%H", "%M", "%S"):
                if len(buf) not in (1, 2):
                    return False, date_str
                date_str = date_str.replace(rep, buf)

            pos += len(buf)
        else:
            # 通常文字
            if pos >= len(value) or value[pos] != ch:
                return False, date_str
            pos += 1
        i += 1

    # 残りの % を初期値で埋める
    now = datetime.now()
    date_str = (date_str.replace("%Y", now.strftime("%Y"))
                        .replace("%m", "1")
                        .replace("%d", "1")
                        .replace("%H", "00")
                        .replace("%M", "00")
                        .replace("%S", "00"))

    # 日付として妥当かチェック
    try:
        datetime.strptime(date_str, "%Y/%m/%d %H:%M:%S")
    except ValueError:
        return False, date_str

    return True, date_str


def format_date(value: str, format_out: str) -> str:
    """
    VBA版 FormatDate のPython版
    """
    if not value:
        return ""

    dt = datetime.strptime(value, "%Y/%m/%d %H:%M:%S")
    rep = {
        "%Y": dt.strftime("%Y"),
        "%y": dt.strftime("%y"),
        "%m": dt.strftime("%m"),
        "%d": dt.strftime("%d"),
        "%H": dt.strftime("%H"),
        "%M": dt.strftime("%M"),
        "%S": dt.strftime("%S"),
    }
    for k, v in rep.items():
        format_out = format_out.replace(k, v)
    return format_out


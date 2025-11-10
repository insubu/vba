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

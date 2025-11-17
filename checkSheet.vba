import os
from datetime import datetime
import pandas as pd


def save_result_to_file(df, MainInfo, AttributeInfo, GetAttributeInfoIndex, PutDQ):
    """
    df: pandas DataFrame，对应 Excel 的 shtTarget 的 UsedRange
    MainInfo: 包含 SaveDirPath / SaveExtension / OriginCell / SaveMode 等信息的对象
    AttributeInfo: 属性定义数组
    GetAttributeInfoIndex: 函数
    PutDQ: 函数（双引号处理）
    """

    try:
        # 保存路径
        save_dir = MainInfo.SaveDirPath
        if not save_dir:
            save_dir = MainInfo.bokTargetPath     # 对应 bokTarget.path

        # 文件名
        file_name = (
            f"{MainInfo.bokTargetBaseName}_"
            f"{datetime.now().strftime('%Y%m%d%H%M%S')}."
            f"{MainInfo.SaveExtension}"
        )
        full_path = os.path.join(save_dir, file_name)

        # 行列范围
        row_min = (
            MainInfo.OriginCell.Row
            if (not MainInfo.OriginCell.AddHeader and
                MainInfo.SaveMode != MainInfo.enumSaveMode.Fixed)
            else MainInfo.OriginCell.Row + 1
        )

        row_max = df.shape[0]      # UsedRange.Rows.Count
        col_max = df.shape[1]      # UsedRange.Columns.Count

        # 打开文件
        with open(full_path, "w", encoding="utf-8") as f:

            # 行循环
            for r in range(row_min - 1, row_max):   # DataFrame 行号从 0 开始
                row_buf = []

                # 列循环
                for c in range(MainInfo.OriginCell.Col - 1, col_max):
                    header = df.iloc[MainInfo.OriginCell.Row - 1, c]
                    data = df.iloc[r, c]

                    if r != MainInfo.OriginCell.Row - 1:
                        # 查属性定义
                        idx = GetAttributeInfoIndex(header)
                        if idx == -1:
                            raise Exception(f"[属性未定义]: {header}")

                        attr = AttributeInfo[idx]
                        if attr.AttrType in [
                            attr.enumAttributeType.Alphanumeric,
                            attr.enumAttributeType.Date,
                            attr.enumAttributeType.Narrow,
                            attr.enumAttributeType.NarrowKana,
                            attr.enumAttributeType.Wide
                        ]:
                            if MainInfo.SaveMode != MainInfo.enumSaveMode.Fixed:
                                data = PutDQ(data)

                    row_buf.append("" if data is None else str(data))

                # 按保存方式插入分隔符
                if MainInfo.SaveMode in (MainInfo.enumSaveMode.Csv,
                                         MainInfo.enumSaveMode.TextComma):
                    sep = ","
                elif MainInfo.SaveMode == MainInfo.enumSaveMode.TextTab:
                    sep = "\t"
                else:
                    sep = ""

                f.write(sep.join(row_buf) + "\n")

        return True

    except Exception as e:
        # 对应 VBA 的 OutputMsg
        print("Error in save_result_to_file:", e)
        return False

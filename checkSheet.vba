import os
from datetime import datetime
import win32com.client as win32


def save_result_to_file(
    shtTarget,
    MainInfo,
    AttributeInfo,
    GetAttributeInfoIndex,
    PutDQ
):
    """
    shtTarget: Excel 工作表对象 (win32com)
    MainInfo: 包含 SaveDirPath / SaveExtension / OriginCell / SaveMode 等
    AttributeInfo: 属性信息数组
    GetAttributeInfoIndex: 函数
    PutDQ: 函数
    """

    try:
        # ===== 保存路径处理 =====
        save_dir = MainInfo.SaveDirPath
        if not save_dir:
            save_dir = MainInfo.bokTargetPath  # 对应 bokTarget.path

        # ===== 文件名 =====
        filename = (
            f"{MainInfo.bokTargetBaseName}_"
            f"{datetime.now().strftime('%Y%m%d%H%M%S')}."
            f"{MainInfo.SaveExtension}"
        )
        full_path = os.path.join(save_dir, filename)

        # ===== UsedRange（对应 VBA）=====
        used = shtTarget.UsedRange
        row_max = used.Rows.Count
        col_max = used.Columns.Count

        # ===== 行起点（与 VBA 完全一致）=====
        if (not MainInfo.OriginCell.AddHeader and
                MainInfo.SaveMode != MainInfo.enumSaveMode.Fixed):
            row_min = MainInfo.OriginCell.Row
        else:
            row_min = MainInfo.OriginCell.Row + 1

        # ===== 打开输出文件 =====
        with open(full_path, "w", encoding="utf-8") as f:

            # ===== 行循环 =====
            for r in range(row_min, row_max + 1):

                buf_list = []

                # ===== 列循环 =====
                for c in range(MainInfo.OriginCell.Col, col_max):

                    header = shtTarget.Cells(MainInfo.OriginCell.Row, c).Value
                    data = shtTarget.Cells(r, c).Value

                    # ----- 非 Header 行处理 -----
                    if r != MainInfo.OriginCell.Row:
                        idx = GetAttributeInfoIndex(header)
                        if idx == -1:
                            raise Exception(f"Attribute undefined: {header}")

                        attr = AttributeInfo[idx]

                        # 数值型以外 → 双引号包围（固定长除外）
                        if attr.AttrType in [
                            attr.enumAttributeType.Alphanumeric,
                            attr.enumAttributeType.Date,
                            attr.enumAttributeType.Narrow,
                            attr.enumAttributeType.NarrowKana,
                            attr.enumAttributeType.Wide
                        ]:
                            if MainInfo.SaveMode != MainInfo.enumSaveMode.Fixed:
                                data = PutDQ(data)

                    buf_list.append("" if data is None else str(data))

                # ===== 分隔符处理 =====
                if MainInfo.SaveMode in (
                    MainInfo.enumSaveMode.Csv,
                    MainInfo.enumSaveMode.TextComma
                ):
                    sep = ","
                elif MainInfo.SaveMode == MainInfo.enumSaveMode.TextTab:
                    sep = "\t"
                else:
                    sep = ""

                f.write(sep.join(buf_list) + "\n")

        return True

    except Exception as e:
        # 对应 VBA ErrHandler
        print("SaveResultToFile Error:", e)
        return False

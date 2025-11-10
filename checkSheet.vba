import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

def check_sheet(file_path: str, main_info, edit_value_func, get_attribute_info_index_func, output_msg_func) -> bool:
    """
    VBAの CheckSheet 関数を Python に移植したもの。

    Args:
        file_path: Excelファイルのパス
        main_info: MainInfo オブジェクト (OriginCell.Row, OriginCell.Col, ResultHeadrName, ErrCellColor, EditCellColorなどを含む)
        edit_value_func: def edit_value(value, attr_index) -> (bool, strErrMsg, strBuf)
        get_attribute_info_index_func: def get_attribute_info_index(header) -> int
        output_msg_func: def output_msg(msg_id, mode, text, icon, app_title)
    """
    try:
        # === Excel ファイルを開く ===
        wb = load_workbook(file_path)
        ws = wb.active

        # UsedRange 相当: pandasで範囲を把握しやすくする
        df = pd.DataFrame(ws.values)
        row_max, col_max = df.shape

        # 結果列を追加
        result_col = col_max + 1
        ws.cell(row=main_info.OriginCell.Row, column=result_col).value = main_info.ResultHeadrName

        err_fill = PatternFill(start_color=main_info.ErrCellColor, end_color=main_info.ErrCellColor, fill_type="solid")
        edit_fill = PatternFill(start_color=main_info.EditCellColor, end_color=main_info.EditCellColor, fill_type="solid")

        err_flag = False
        data_number = 0

        # === 行ループ ===
        for row in range(main_info.OriginCell.Row + 1, row_max + 1):
            data_number += 1
            print(f"{main_info.AppTitle} 処理中... [{data_number}/{row_max - main_info.OriginCell.Row}]")

            # === 列ループ ===
            for col in range(main_info.OriginCell.Col, col_max + 1):
                header = ws.cell(row=main_info.OriginCell.Row, column=col).value
                data = ws.cell(row=row, column=col).value

                # 属性定義の確認
                attr_index = get_attribute_info_index_func(header)
                if attr_index == -1:
                    output_msg_func("MSG_104", "MODE_DLG", header, "EXCLAMATION", main_info.AppTitle)
                    wb.save(file_path)
                    return False

                # 値の編集・チェック
                ok, err_msg, buf = edit_value_func(data, attr_index)
                if not ok:
                    ws.cell(row=row, column=col).fill = err_fill
                    ws.cell(row=row, column=result_col).value = f"NG [{header}:{err_msg}]"
                    err_flag = True
                    break  # 次の行へ

                if data != buf:
                    ws.cell(row=row, column=col).fill = edit_fill
                    ws.cell(row=row, column=col).value = buf

            else:
                # 全列OK
                ws.cell(row=row, column=result_col).value = "OK"

        # === 結果 ===
        wb.save(file_path)

        if err_flag:
            output_msg_func("MSG_201", "MODE_DLG", "", "EXCLAMATION", main_info.AppTitle)
            return False
        return True

    except Exception as e:
        output_msg_func("MSG_999", "MODE_ALL", f"CheckSheet#Error#{type(e).__name__}#{str(e)}", "CRITICAL", main_info.AppTitle)
        return False

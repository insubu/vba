import win32com.client as win32

def check_sheet(main_info, attribute_info, make_worksheet_func, get_attr_index_func, edit_value_func):
    """
    main_info: object with OriginCell, ResultHeaderName, ErrCellColor, EditCellColor, etc.
    attribute_info: list of attribute objects with AttrName
    make_worksheet_func: callable that prepares the work sheet (returns True/False)
    get_attr_index_func: callable(header_name) -> int
    edit_value_func: callable(data_str, attr_index) -> (ok:bool, err_msg:str, edited_value:str)
    """
    excel = win32.Dispatch("Excel.Application")
    app = excel.Application

    try:
        app.Cursor = -4143  # xlWait
        app.ScreenUpdating = False

        # --- make worksheet ---
        if not make_worksheet_func():
            return False

        # assuming shtTarget and bokTarget are global or accessible somehow
        wb_target = app.ActiveWorkbook
        ws_target = wb_target.ActiveSheet

        # --- get used range bounds ---
        used_range = ws_target.UsedRange
        lngRowMax = used_range.Rows.Count
        lngColMax = used_range.Columns.Count

        # --- create result column ---
        lngColResult = lngColMax + 1
        ws_target.Cells(main_info.OriginCell.Row, lngColResult).Value = main_info.ResultHeaderName

        blnErrFlag = False
        lngDataNumber = 0

        # --- iterate rows ---
        for lngRow in range(main_info.OriginCell.Row + 1, lngRowMax + 1):
            lngDataNumber += 1
            app.StatusBar = f"{main_info.AppTitle} 処理中です... [{lngDataNumber}/{lngRowMax - main_info.OriginCell.Row}件]"

            # --- iterate columns ---
            for lngCol in range(main_info.OriginCell.Col, lngColMax + 1):
                strHeader = ws_target.Cells(main_info.OriginCell.Row, lngCol).Value
                strData = ws_target.Cells(lngRow, lngCol).Value

                # attribute index
                lngAttrIndex = get_attr_index_func(strHeader)
                if lngAttrIndex == -1:
                    # 未定義属性エラー
                    msg = f"属性未定義: {strHeader}"
                    print(msg)  # or call OutputMsg(MSG_104, ...)
                    return False

                ok, err_msg, edited_value = edit_value_func(strData, lngAttrIndex)

                if not ok:
                    # 編集エラー
                    ws_target.Cells(lngRow, lngCol).Interior.ColorIndex = main_info.ErrCellColor
                    ws_target.Cells(lngRow, lngColResult).Value = f"NG [{strHeader}:{err_msg}]"
                    blnErrFlag = True
                    break  # go to next row

                if strData != edited_value:
                    ws_target.Cells(lngRow, lngCol).Interior.ColorIndex = main_info.EditCellColor
                    ws_target.Cells(lngRow, lngCol).Value = edited_value

            # --- if all columns done ---
            if lngCol == lngColMax + 1:
                ws_target.Cells(lngRow, lngColResult).Value = "OK"

        app.StatusBar = False
        app.ScreenUpdating = True

        # --- error summary dialog ---
        if blnErrFlag:
            print("エラーが存在します。確認してください。")
            return False

        return True

    except Exception as e:
        # error handler
        print(f"[Error] CheckSheet: {e}")
        return False

    finally:
        app.Cursor = -4143  # xlDefault
        app.ScreenUpdating = True
        app.StatusBar = False

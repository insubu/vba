import win32com.client as win32

def make_worksheet(main_info, attribute_info, work_tag="_WORK"):
    """
    main_info: object with OriginCell, SaveDirPath, TrimHeaderCrLf, TrimHeaderSpace, etc.
    attribute_info: list of objects/dicts with AttrName, ColPos
    """
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Optional: True to see the operations
    
    # --- 1. Get active workbook ---
    if excel.Workbooks.Count <= 1:
        raise RuntimeError("Only one workbook open, cannot process.")
    
    wb_target = excel.ActiveWorkbook
    
    # --- 2. Check save path ---
    if not wb_target.Path and not getattr(main_info, "SaveDirPath", ""):
        raise RuntimeError("No save path defined.")
    
    # --- 3. Get active sheet ---
    ws_target = wb_target.ActiveSheet
    work_name = f"{ws_target.Name}{work_tag}"
    
    # --- 4. Delete existing work sheet if exists ---
    for ws in wb_target.Worksheets:
        if ws.Name == work_name:
            response = excel.Application.InputBox(
                "Work sheet exists. OK to recreate?", "Confirm", Type=1
            )
            # Cancel (or No) behavior can be added if needed
            ws.Delete()
            break
    
    # --- 5. Copy sheet ---
    ws_target.Copy(After=wb_target.Sheets(wb_target.Sheets.Count))
    ws_work = wb_target.Sheets(wb_target.Sheets.Count)
    ws_work.Name = work_name
    
    # --- 6. Insert header row if needed ---
    if main_info.OriginCell.AddHeader:
        row = main_info.OriginCell.Row
        col = main_info.OriginCell.Col
        ws_work.Rows(row).Insert()
        for attr in attribute_info:
            ws_work.Cells(row, col + attr.ColPos - 1).Value = attr.AttrName
    
    # --- 7. Delete upper rows / left columns if needed ---
    if main_info.OriginCell.DeleteUpperRow:
        if main_info.OriginCell.Row > 1:
            ws_work.Range(ws_work.Rows(1), ws_work.Rows(main_info.OriginCell.Row - 1)).Delete()
            main_info.OriginCell.Row = 1
        if main_info.OriginCell.Col > 1:
            ws_work.Range(ws_work.Columns(1), ws_work.Columns(main_info.OriginCell.Col - 1)).Delete()
            main_info.OriginCell.Col = 1
    
    # --- 8. Trim header row ---
    row = main_info.OriginCell.Row
    max_col = ws_work.UsedRange.Columns.Count
    
    for col in range(main_info.OriginCell.Col, max_col + 1):
        val = ws_work.Cells(row, col).Value
        if isinstance(val, str):
            # Remove line breaks
            if main_info.TrimHeaderCrLf:
                val = val.replace("\r", "").replace("\n", "")
            # Remove spaces
            trim_mode = getattr(main_info, "TrimHeaderSpace", "TrimBoth")
            if trim_mode == "TrimAll":
                val = val.replace(" ", "")
            elif trim_mode == "TrimBoth":
                val = val.strip()
            elif trim_mode == "TrimLeft":
                val = val.lstrip()
            elif trim_mode == "TrimRight":
                val = val.rstrip()
            ws_work.Cells(row, col).Value = val
    
    # --- Optional: Save workbook ---
    # wb_target.Save()  # or SaveAs to new path
    
    return True

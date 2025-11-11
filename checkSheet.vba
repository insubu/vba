def edit_value(str_value: str, attr_index: int, attribute_info, replace_info):
    """
    Python version of VBA EditValue
    Params:
        str_value (str): 原值
        attr_index (int): AttributeInfo 的索引
        attribute_info (list[dict]): 属性定义信息（取代 VBA 的 AttributeInfo 数组）
        replace_info (list[dict]): 替换规则（取代 VBA 的 ReplaceInfo）
    Returns:
        (bool, str, str): (成功/失败, 错误信息, 编辑后的值)
    """

    try:
        attr = attribute_info[attr_index]
        buf = str_value or ""
        err_msg = ""
        edit_value = ""

        # --- 改行削除 ---
        if attr.get("TrimCrLf"):
            buf = buf.replace("\r", "").replace("\n", "")

        # --- スペース削除 ---
        trim_mode = attr.get("TrimSpace")
        if trim_mode == "TrimAll":
            buf = buf.replace(" ", "")
        elif trim_mode == "TrimBoth":
            buf = buf.strip()
        elif trim_mode == "TrimLeft":
            buf = buf.lstrip()
        elif trim_mode == "TrimRight":
            buf = buf.rstrip()

        # --- 必須チェック ---
        if attr.get("Indispensable") and buf == "":
            return False, "必須属性が未入力です。", ""

        # --- 文字コードチェック ---
        for ch in buf:
            if not is_permitted_code(ch):
                return False, f"使用不可文字[{ch}]が入力されています。", ""

        # --- 属性の型による文字変換 ---
        attr_type = attr.get("AttrType")
        if attr_type == "IntegerNumber":
            buf = to_half_width(buf)
            if buf and (not buf.isdigit() or "," in buf or "." in buf):
                return False, "数値以外または小数点/カンマが入力されています。", ""
        elif attr_type == "SmallNumber":
            buf = to_half_width(buf)
            if buf and (not is_numeric(buf) or "," in buf):
                return False, "数値以外が入力されています。", ""
        elif attr_type == "Narrow":
            buf = to_half_width(buf)
            for ch in buf:
                if not is_narrow(ch):
                    return False, "半角対象文字以外が入力されています。", ""
        elif attr_type == "Date":
            date_str = get_date_str(buf, attr.get("DateFormat_In"))
            if not date_str:
                return False, "入力された日付の書式が不正です。", ""
        else:
            # --- 置換処理 ---
            replaced = False
            for r in replace_info:
                if r["ReplaceMode"] == "Complete" and buf == r["KeyString"]:
                    buf = r["ReplaceString"]
                    replaced = True
                    break

            if not replaced:
                work = ""
                i = 0
                while i < len(buf):
                    ch = buf[i]
                    matched = False
                    for r in replace_info:
                        if r["ReplaceMode"] == "Partial" and buf.startswith(r["KeyString"], i):
                            work += r["ReplaceString"]
                            i += len(r["KeyString"])
                            matched = True
                            break
                    if not matched:
                        if attr_type == "Wide":
                            ch = to_wide(ch)
                        elif attr_type == "Alphanumeric":
                            ch = to_half_width(ch) if is_alphanumeric(ch) else to_wide(ch)
                        elif attr_type == "NarrowKana":
                            ch = narrow_kana_to_wide(ch)
                        work += ch
                        i += 1
                buf = work

        # --- バイト数加工 ---
        byte_mode = attr.get("ByteEditMode")
        left_size = attr.get("ByteSize_Left")
        right_size = attr.get("ByteSize_Right")

        if byte_mode == "Fixed":
            if not is_complete_byte(buf, left_size, right_size):
                return False, "入力された文字のバイト数が規定値と異なります。", ""
        elif byte_mode == "Complete":
            if not is_permitted_byte(buf, left_size, right_size):
                return False, "入力された文字のバイト数が規定値を超えています。", ""
            if attr_type == "Date":
                buf = format_date(date_str, attr.get("DateFormat_Out"))
            else:
                buf = fill_string(buf, left_size, right_size, attr.get("CompleteChar"))
        elif byte_mode == "Max":
            if not is_permitted_byte(buf, left_size, right_size):
                return False, "入力された文字のバイト数が規定値を超えています。", ""

        # --- 大文字小文字統一 ---
        letter_type = attr.get("LetterType")
        if letter_type == "Capital":
            buf = buf.upper()
        elif letter_type == "Small":
            buf = buf.lower()

        return True, "", buf

    except Exception as e:
        # VBA の ErrHandler 相当
        print(f"[Error] EditValue: {e}")
        return False, str(e), ""

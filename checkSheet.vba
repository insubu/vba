from enum import Enum

class ReplaceMode(Enum):
    Complete = 0    # 完全一致
    Partial = 1     # 文字列一致


class ReplaceInfoItem:
    def __init__(self, key_string="", replace_string="", replace_mode=None):
        self.KeyString = key_string
        self.ReplaceString = replace_string
        self.ReplaceMode = replace_mode


def read_replace_sheet(sheet):
    """
    置換シートを読み込む
    Returns: bool
    """
    replace_info_list = []
    try:
        # used range rows count equivalent
        row_max = sheet.used_range_last_row()  # ← implement this for your environment

        for row in range(3, row_max + 1):
            # === 変換前 ===
            key_col = "変換前"
            key_str = read_csv_sheet(sheet, key_col, row)
            if not key_str:
                # 未入力の場合は有効行でないと判断
                break

            # === 変換後 ===
            replace_col = "変換後"
            replace_str = read_csv_sheet(sheet, replace_col, row)
            if not replace_str:
                show_ini_sheet(sheet)
                output_msg("MSG_002", "MODE_ALL", f"{sheet.name}#{replace_col}#{row}")
                return False

            # === 完全一致 / 文字列一致 ===
            match_col = "完全一致"
            mode_str = read_csv_sheet(sheet, match_col, row).upper()
            if mode_str == "完全一致":
                mode = ReplaceMode.Complete
            elif mode_str == "文字列一致":
                mode = ReplaceMode.Partial
            else:
                show_ini_sheet(sheet)
                output_msg("MSG_002", "MODE_ALL", f"{sheet.name}#{match_col}#{row}")
                return False

            new_item = ReplaceInfoItem(key_str, replace_str, mode)

            # === 整合性チェック ===
            for existing in replace_info_list:
                # [変換前]重複チェック
                if (existing.ReplaceMode == ReplaceMode.Partial or
                    new_item.ReplaceMode == ReplaceMode.Partial):
                    if (existing.KeyString in new_item.KeyString or
                        new_item.KeyString in existing.KeyString):
                        show_ini_sheet(sheet)
                        output_msg("MSG_003", "MODE_ALL",
                                   f"{sheet.name}#変換前#{row}")
                        return False

                # 循環参照チェック
                if (new_item.ReplaceString == existing.KeyString and
                    new_item.KeyString == existing.ReplaceString):
                    show_ini_sheet(sheet)
                    output_msg("MSG_004", "MODE_ALL",
                               f"{sheet.name}#変換後#{row}")
                    return False

            replace_info_list.append(new_item)

        return True

    except Exception as e:
        # corresponds to ErrHandler
        output_msg("MSG_999", "MODE_ALL",
                   f"ReadReplaceSheet#{type(e).__name__}#{str(e)}")
        return False

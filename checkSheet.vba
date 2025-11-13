def parse_byte_edit_mode(attr: AttributeSheet, buf: str):
    if buf == "固定":
        attr.byte_edit_mode = ByteEditMode.FIXED
    elif buf == "最大":
        attr.byte_edit_mode = ByteEditMode.MAX
    elif buf.startswith("補完:"):
        attr.byte_edit_mode = ByteEditMode.COMPLETE
        param = buf.split(":", 1)[1]
        if attr.attr_type == AttributeType.DATE:
            attr.date_format_out = param
        elif attr.attr_type in (AttributeType.INTEGER_NUMBER, AttributeType.SMALL_NUMBER):
            if param in ("0", " "):
                attr.complete_char = param
            else:
                raise ValueError(f"Invalid 補完 param for numeric type: {param}")
        else:
            if len(param) != 1:
                raise ValueError(f"Invalid 補完 char '{param}'")
            attr.complete_char = param
    else:
        raise ValueError(f"Invalid バイト数加工 '{buf}'")

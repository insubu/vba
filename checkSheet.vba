def parse_byte_edit_mode(attr: AttributeSheet, buf: str):
    if buf == "固定":
        attr.ByteEditMode = ByteEditMode.Fixed
    elif buf == "最大":
        attr.ByteEditMode = ByteEditMode.Max
    elif buf.startswith("補完:"):
        attr.ByteEditMode = ByteEditMode.Complete
        param = buf.split(":", 1)[1]
        if attr.AttrType == AttributeType.Date:
            attr.DateFormat_Out = param
        elif attr.AttrType in (AttributeType.IntegerNumber, AttributeType.SmallNumber):
            if param in ("0", " "):
                attr.CompleteChar = param
            else:
                raise ValueError(f"Invalid 補完 param for numeric type: {param}")
        else:
            if len(param) != 1:
                raise ValueError(f"Invalid 補完 char '{param}'")
            attr.CompleteChar = param
    else:
        raise ValueError(f"Invalid バイト数加工 '{buf}'")

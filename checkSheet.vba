# -*- coding: utf-8 -*-
# Equivalent of VBA: IsPermittedCode / IsPermittedByte

class PeriodCode:
    def __init__(self, code_from, code_to):
        self.code_from = code_from  # e.g. "3000"
        self.code_to = code_to      # e.g. "30FF"

class MainInfo:
    def __init__(self):
        self.PeriodCode = []        # list of PeriodCode objects
        self.PeriodCodeCount = 0

# Example global structure
main_info = MainInfo()
# main_info.PeriodCode = [PeriodCode("3000", "30FF"), PeriodCode("4E00", "9FFF")]
# main_info.PeriodCodeCount = len(main_info.PeriodCode)


# ------------------------------
# Check if a character code is permitted
# ------------------------------
def is_permitted_code(char: str, main_info: MainInfo) -> bool:
    """Check if given character is within permitted code ranges."""
    if not char:
        return False

    code_point = ord(char)

    # ASCII range 0–255
    if 0 <= code_point <= 255:
        return True

    # Otherwise check MainInfo.PeriodCode ranges
    for i in range(main_info.PeriodCodeCount):
        pc = main_info.PeriodCode[i]
        from_code = int(pc.code_from, 16)
        to_code = int(pc.code_to, 16)
        if from_code <= code_point <= to_code:
            return True

    return False


# ------------------------------
# Check if a string’s byte length is within allowed limits
# ------------------------------
def is_permitted_byte(value: str, byte_left: int, byte_right: int) -> bool:
    """Check if a numeric string fits in allowed byte lengths (Shift-JIS)."""
    value_left = ""
    value_right = ""

    if byte_right != 0:
        # Fractional number case
        if "." in value:
            value_left, value_right = value.split(".", 1)
        else:
            value_left = value
    else:
        # Non-decimal
        value_left = value

    # Check integer part
    if byte_left != 0:
        if len(value_left.encode("shift_jis")) > byte_left:
            return False

    # Check fractional part
    if byte_right != 0:
        if len(value_right.encode("shift_jis")) > byte_right:
            return False

    return True

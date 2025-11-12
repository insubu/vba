import jaconv

def is_alphanumeric(char: str) -> bool:
    """
    Check if a given character is alphanumeric (A–Z, a–z, 0–9).
    Equivalent to VBA IsAlphanumeric with StrConv(..., vbNarrow).
    """
    if not char:
        return False

    # Convert to half-width (vbNarrow equivalent)
    char_half = jaconv.z2h(char, kana=False, digit=True, ascii=True)

    code = ord(char_half)
    # 0-9: 48–57, A-Z: 65–90, a-z: 97–122
    if (48 <= code <= 57) or (65 <= code <= 90) or (97 <= code <= 122):
        return True
    return False

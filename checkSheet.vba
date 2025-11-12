def is_narrow(char: str) -> bool:
    """
    Check if a given character is half-width (ASCII range 32–126).
    Equivalent to VBA IsNarrow().
    """
    if not char:
        return False

    code = ord(char)
    if 32 <= code <= 126:
        return True
    return False

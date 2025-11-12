from datetime import datetime

def format_date(str_value: str, str_format_out: str) -> str:
    """
    Convert a date string to a specified format, 
    mimicking VBA FormatDate() logic.
    """
    if not str_value:
        # Equivalent to: If strValue = "" Then GoTo EndHandler
        return ""

    try:
        # Parse input date string flexibly
        dt_value = datetime.fromisoformat(str_value)
    except ValueError:
        # Try common formats if not ISO
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
            try:
                dt_value = datetime.strptime(str_value, fmt)
                break
            except ValueError:
                continue
        else:
            # If all parsing fails
            return ""

    # Start with output template
    buf = str_format_out

    # Replace VBA-style placeholders (%Y, %m, etc.)
    buf = buf.replace("%Y", dt_value.strftime("%Y"))
    buf = buf.replace("%y", dt_value.strftime("%y"))
    buf = buf.replace("%m", dt_value.strftime("%m"))
    buf = buf.replace("%d", dt_value.strftime("%d"))
    buf = buf.replace("%H", dt_value.strftime("%H"))  # 24-hour format
    buf = buf.replace("%M", dt_value.strftime("%M"))
    buf = buf.replace("%S", dt_value.strftime("%S"))

    return buf

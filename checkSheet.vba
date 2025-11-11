from enum import Enum

class AttributeType(Enum):
    Non = 0
    Narrow = 1
    Wide = 2
    Alphanumeric = 3
    NarrowKana = 4
    IntegerNumber = 5
    SmallNumber = 6
    Date = 7

def process_attribute(attribute_info, attribute_index, sheet, row):
    key = "型"
    buf = read_csv_sheet(sheet, key, row)

    if buf == "":
        attribute_info[attribute_index].AttrType = AttributeType.Non
    elif buf == "半角":
        attribute_info[attribute_index].AttrType = AttributeType.Narrow
    elif buf == "全角":
        attribute_info[attribute_index].AttrType = AttributeType.Wide
    elif buf == "英数字":
        attribute_info[attribute_index].AttrType = AttributeType.Alphanumeric
    elif buf == "半角カナ":
        attribute_info[attribute_index].AttrType = AttributeType.NarrowKana
    elif buf == "整数":
        attribute_info[attribute_index].AttrType = AttributeType.IntegerNumber
    elif buf == "小数":
        attribute_info[attribute_index].AttrType = AttributeType.SmallNumber
    elif buf.startswith("日付:"):
        attribute_info[attribute_index].AttrType = AttributeType.Date
        # extract date format after ":"
        attribute_info[attribute_index].DateFormat_In = buf.split(":", 1)[1]
    else:
        raise ValueError(f"Invalid attribute type: {buf}")

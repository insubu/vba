Public Enum enumSaveMode
    Csv                 'CSV
    TextTab             'TEXT（タブ）
    TextComma           'TEXT（カンマ）
    Fixed               '固定長
End Enum


class SaveMode(Enum):
    Csv = 0
    TextTab = 1
    TextComma = 2
    Fixed = 3

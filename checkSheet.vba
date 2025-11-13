例外が発生しました: UnicodeEncodeError
'shift_jis' codec can't encode character '\uff02' in position 7: illegal multibyte sequence
  File "C:\Users\fxYb316.DC00\Downloads\CheckCharTool\baseMain.py", line 363, in is_permitted_byte
    if len(value_left.encode("shift_jis")) > byte_left:
           ~~~~~~~~~~~~~~~~~^^^^^^^^^^^^^
  File "C:\Users\fxYb316.DC00\Downloads\CheckCharTool\baseMain.py", line 298, in edit_value
    if not is_permitted_byte(buf, left_size, right_size):
           ~~~~~~~~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\fxYb316.DC00\Downloads\CheckCharTool\baseMain.py", line 146, in check_sheet
    ok, err_msg, edited_value = edit_value(strData, lngAttrIndex)
                                ~~~~~~~~~~^^^^^^^^^^^^^^^^^^^^^^^
  File "C:\Users\fxYb316.DC00\Downloads\CheckCharTool\tst.py", line 30, in <module>
    baseMain.check_sheet()
    ~~~~~~~~~~~~~~~~~~~~^^
UnicodeEncodeError: 'shift_jis' codec can't encode character '\uff02' in position 7: illegal multibyte sequence

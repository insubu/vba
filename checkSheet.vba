import os
from datetime import datetime
import win32com.client as win32


def save_result_to_file(
    shtTarget,
    MainInfo,
    AttributeInfo,
    GetAttributeInfoIndex,
    PutDQ
):
    """
    VBA の SaveResultToFile と等価の処理を Python で実装
    shtTarget: Excel シート (win32com)
    """

    try:
        # ===== 保存先ディレクトリの決定 =====
        # MainInfo に指定がなければ、対象ブックと同じパスを使用
        save_dir = MainInfo.SaveDirPath
        if not save_dir:
            save_dir = MainInfo.bokTargetPath

        # ===== 出力ファイル名を作成（VBA と同じ YYYYMMDDhhmmss 形式）=====
        filename = (
            f"{MainInfo.bokTargetBaseName}_"
            f"{datetime.now().strftime('%Y%m%d%H%M%S')}."
            f"{MainInfo.SaveExtension}"
        )
        full_path = os.path.join(save_dir, filename)

        # ===== UsedRange を使用して最終行・最終列を取得（VBA と同等）=====
        used = shtTarget.UsedRange
        row_max = used.Rows.Count
        col_max = used.Columns.Count

        # ===== ヘッダの出力行の開始位置を決定 =====
        # AddHeader = False かつ 固定長以外 → ヘッダ行も含めて処理
        if (not MainInfo.OriginCell.AddHeader and
                MainInfo.SaveMode != MainInfo.enumSaveMode.Fixed):
            row_min = MainInfo.OriginCell.Row
        else:
            # それ以外はヘッダ行をスキップ（データのみ）
            row_min = MainInfo.OriginCell.Row + 1

        # ===== 出力ファイルを開く =====
        with open(full_path, "w", encoding="utf-8") as f:

            # ===== 行ループ（OriginCell.Row から最終行まで）=====
            for r in range(row_min, row_max + 1):

                buf_list = []

                # ===== 列ループ =====
                for c in range(MainInfo.OriginCell.Col, col_max):

                    # ヘッダ名（属性名）
                    header = shtTarget.Cells(MainInfo.OriginCell.Row, c).Value
                    # データ本体
                    data = shtTarget.Cells(r, c).Value

                    # ---- ヘッダ行以外の処理 ----
                    if r != MainInfo.OriginCell.Row:

                        # 属性名から AttributeInfo のインデックス取得
                        idx = GetAttributeInfoIndex(header)
                        if idx == -1:
                            raise Exception(f"未定義の属性です: {header}")

                        attr = AttributeInfo[idx]

                        # 数値型以外 → 固定長でなければダブルクォートで囲む
                        if attr.AttrType in [
                            attr.enumAttributeType.Alphanumeric,
                            attr.enumAttributeType.Date,
                            attr.enumAttributeType.Narrow,
                            attr.enumAttributeType.NarrowKana,
                            attr.enumAttributeType.Wide
                        ]:
                            if MainInfo.SaveMode != MainInfo.enumSaveMode.Fixed:
                                data = PutDQ(data)

                    # セルが None の場合は空文字にする
                    buf_list.append("" if data is None else str(data))

                # ===== 保存モードに応じた区切り文字を決定 =====
                if MainInfo.SaveMode in (
                    MainInfo.enumSaveMode.Csv,
                    MainInfo.enumSaveMode.TextComma
                ):
                    sep = ","
                elif MainInfo.SaveMode == MainInfo.enumSaveMode.TextTab:
                    sep = "\t"
                else:
                    sep = ""

                # 1 行分を出力
                f.write(sep.join(buf_list) + "\n")

        return True

    except Exception as e:
        # VBA の ErrHandler 相当
        print("SaveResultToFile エラー:", e)
        return False

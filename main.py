from pathlib import Path

import win32com.client

FINISHED_FILE_SUFFIX = "_脚注付"


def load_excel(excel_path: Path) -> dict[str, str]:
    """
    Excelファイルから脚注マッピングを読み込む
    1列目: 脚注番号、2列目: 脚注テキスト
    """
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    try:
        workbook = excel.Workbooks.Open(str(excel_path), ReadOnly=True)
        sheet = workbook.Worksheets(1)  # 最初のシートを使用

        # 使用範囲を取得
        used_range = sheet.UsedRange
        rows = used_range.Rows.Count

        d = {}
        for i in range(1, rows + 1):
            key = sheet.Cells(i, 1).Value
            value = sheet.Cells(i, 2).Value

            if key is not None and value is not None:
                # 文字列化して前後の空白を削除
                d[str(int(key)).strip()] = str(value).strip()

        workbook.Close(SaveChanges=False)
        return d

    finally:
        excel.Quit()


def apply_footnotes(docx_path: Path) -> None:
    excel_path = docx_path.with_suffix(".xlsx")
    if not excel_path.exists():
        # .xlsxが無ければ.xlsも試す
        excel_path = docx_path.with_suffix(".xls")
        if not excel_path.exists():
            print(
                f"{docx_path.name}の脚注定義ファイル {excel_path.name} がありません。スキップします。"
            )
            return

    footnote_mapping = load_excel(excel_path)

    output_path = docx_path.with_name(f"{docx_path.stem}{FINISHED_FILE_SUFFIX}.docx")

    # Wordアプリケーションを起動
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        # ドキュメントを開く
        doc = word.Documents.Open(str(docx_path))

        # 本文のRangeを使って検索
        rng = doc.Content
        find = rng.Find
        find.ClearFormatting()

        # ワイルドカード検索を有効化
        find.MatchWildcards = True

        # マーカーを検索して脚注を追加
        while find.Execute(FindText="【【[0-9]{1,}】】", Forward=True):
            # 数字部分
            key = str(rng.Text).replace("【", "").replace("】", "")
            # マーカーを削除
            rng.Delete()

            # 脚注を追加
            footnote = doc.Footnotes.Add(Range=rng)
            footnote_text = footnote_mapping.get(key, None)
            if footnote_text is None:
                raise KeyError(
                    f"脚注番号 {key} に対応する脚注がExcel内に記載されていません。"
                )
            footnote.Range.Text = footnote_text

        # 別名で保存
        doc.SaveAs2(str(output_path))

        print(f"==> 処理完了: {output_path.name}")
        print(f"==> 追加した脚注数: {doc.Footnotes.Count}")

        doc.Close()

    except Exception as e:
        print(f"エラーが発生しました: {e}")
        raise

    finally:
        # Wordを終了
        word.Quit()


if __name__ == "__main__":
    data_dir = Path(__file__).parent / "data"

    for c in data_dir.iterdir():
        if c.is_dir():
            continue
        if c.name.startswith("~$"):
            continue
        if c.suffix == ".docx" and not c.stem.endswith(FINISHED_FILE_SUFFIX):
            print(f"処理中：{c.name}")
            apply_footnotes(c)

    input("\n\n処理が完了しました！\n何かキーを押すと終了します")

from pathlib import Path

import win32com.client

FINISHED_FILE_SUFFIX = "_脚注付"
FMISSING_FILE_SUFFIX = "_欠番"


def find_missing_numbers(numbers: list[int]) -> list[int]:
    """
    リスト内の欠番を検出する。
    連続する整数列（min〜max）を期待値として、存在しない値を返す。
    """
    if not numbers:
        return []

    number_set = set(numbers)
    full_range = range(min(numbers), max(numbers) + 1)

    return [n for n in full_range if n not in number_set]


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

        # 注番号を控えていく
        footnote_nums: list[int] = []

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
            footnote_nums.append(int(key))

        # 別名で保存
        doc.SaveAs2(str(output_path))

        print(f"==> 処理完了: {output_path.name}")
        print(f"==> 追加した脚注数: {doc.Footnotes.Count}")

        missing_nums = find_missing_numbers(footnote_nums)
        if 0 < len(missing_nums):
            print(
                f"==> 注番号の抜けが{len(missing_nums)}件ありました。ファイルに出力します"
            )
            skipped = [
                (num, footnote_mapping.get(str(num), "（原稿中になし）"))
                for num in missing_nums
            ]
            out_missing_path = output_path.with_name(
                f"{output_path.stem}{FMISSING_FILE_SUFFIX}.txt"
            )
            out_missing_path.write_text(
                "\n".join([f"■{s[0]}\n{s[1]}\n" for s in skipped]), encoding="utf-8"
            )

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

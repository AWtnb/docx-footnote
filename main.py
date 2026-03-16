import csv
from pathlib import Path

import win32com.client

FINISHED_FILE_SUFFIX = "_脚注付"


def load_csv(csv_path: Path) -> dict[str, str]:
    d = {}
    with open(csv_path, encoding="utf-8") as f:
        reader = csv.reader(f)
        for line in reader:
            row = [s.strip() for s in line]
            d[row[0]] = row[1]
    return d


def apply_footnotes(docx_path: Path) -> None:
    csv_path = c.with_suffix(".csv")
    if not csv_path.exists():
        print(
            f"{c.name}の脚注定義ファイル {csv_path.name} がありません。スキップします。"
        )
        return

    footnote_mapping = load_csv(csv_path)

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
                    f"脚注番号 {key} に対応する脚注がCSV内に記載されていません。"
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
    data_ditr = Path(__file__).parent / "data"

    for c in data_ditr.iterdir():
        if c.is_dir():
            continue
        if c.name.startswith("~$"):
            continue
        if c.suffix == ".docx" and not c.stem.endswith(FINISHED_FILE_SUFFIX):
            print(f"処理中：{c.name}")
            apply_footnotes(c)

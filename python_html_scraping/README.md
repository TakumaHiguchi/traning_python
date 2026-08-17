# 名言サイト(quotes.toscrape.com) スクレイピングツール

スクレイピング練習用サイト [quotes.toscrape.com](https://quotes.toscrape.com/) から名言・著者・タグを取得し、Excelファイルに自動で蓄積していくPythonスクリプトです。

## できること

- 対象ページのHTMLを取得し、`BeautifulSoup`で解析して「名言」「著者」「タグ」を抽出
- 取得結果をExcelファイル(`quote_list.xlsx`)に保存
- 実行するたびに、既存データへ新しい名言を追記（同じ名言の重複は自動で排除）
- 見出し行の太字・列幅を内容に応じて自動調整
- サイトへのアクセス失敗・該当データが0件だった場合も、安全に処理を終了
- 出力先のExcelファイルが開いたまま実行された場合も、エラーで落ちずに分かりやすいメッセージを表示

## 使用技術

- Python 3
- [requests](https://requests.readthedocs.io/) - Webページの取得
- [BeautifulSoup(bs4)](https://www.crummy.com/software/BeautifulSoup/) - HTMLの解析
- [pandas](https://pandas.pydata.org/) - データの整形・重複排除
- [openpyxl](https://openpyxl.readthedocs.io/) - Excelファイルの書き込み・装飾

## 使い方

### 1. 必要なライブラリをインストール

```bash
pip install requests beautifulsoup4 pandas openpyxl
```

### 2. 実行

```bash
python scraping.py
```

実行すると、スクリプトと同じフォルダに `quote_list.xlsx` が作成（既にあれば更新）されます。

## 工夫した点

- HTML解析はページ全体から一気に検索するのではなく、名言ブロック(`div.quote`)ごとに絞り込んでから著者・タグを検索することで、複数件のデータがズレて対応してしまう事故を防いでいます
- サイトへのアクセス失敗、該当データ0件、Excelファイルが開いていて保存できない、といった実行時に起こりうるトラブルを想定し、それぞれ安全にメッセージを出して終了するようにしています
- 日本語(全角文字)を含む文章の見た目の幅を考慮して、Excelの列幅を内容に応じて自動計算しています

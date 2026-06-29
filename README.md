# 韓国語-漢字変換ツール ドキュメント

## 概要

このツールは、Word 文書（`.docx`）内の韓国語表記を Excel 辞書に基づいて漢字へ変換するためのツールです。

主な処理は以下です。

- Excel 辞書から「韓国語 -> 漢字」の対応表を読み込む
- Excel 辞書から「旧字体 -> 新字体」の対応表を読み込む
- Word 本文、ヘッダー、フッターの文字列を変換する
- 段落をまたいで分割された語も可能な範囲で変換する
- フォントを `Meiryo UI` に統一する
- `()` または `（）` 内の文字幅を 95% に圧縮する
- 変換後の `.docx` を保存し、可能であれば PDF も出力する

## フォルダ構成

```text
.
+-- README.md
+-- ConvertAllText_ハンドオーバー資料_2.docx
+-- 変換ツール引き継ぎ.docx
+-- new
|   +-- ConvertAllText_improved.xlsm
|   +-- convert_hanja.py
|   +-- convert_hanja_safe.py
|   +-- convert_hanja_claude.py
|   +-- convert_hanjaorg.py
|   +-- convert_hanja.bat
|   +-- 実行_変換.bat
|   +-- 実行_変換Date.bat
+-- old
    +-- ConvertAllText_Copilot.xlsm
    +-- ConvertAllText_improved.xlsm
    +-- PDF
```

## 主要ファイル

### `new/convert_hanja.py`

現行の主実装です。

Word 文書を `python-docx` で読み込み、内部の WordprocessingML（OOXML）を直接処理します。辞書置換、フォント統一、括弧内文字幅調整、PDF 変換までを担当します。

### `new/ConvertAllText_improved.xlsm`

変換辞書です。拡張子は `.xlsm` ですが、現状 VBA マクロは含まれておらず、辞書データとして利用されています。

- `Sheet1`: 韓国語から漢字への変換辞書
  - A 列: `Korean`
  - C 列: `Hanja`
- `Sheet2`: 旧字体から新字体への変換表
  - A 列: `旧字体`
  - B 列: `新字体`

### `new/実行_変換Date.bat`

通常運用向けの起動ファイルです。

`.docx` ファイルをこの bat にドラッグ＆ドロップすると、`new/output` フォルダへタイムスタンプ付きで変換後ファイルを出力します。

出力例:

```text
new/output/入力ファイル名_hanja_YYYYMMDD_HHMM.docx
new/output/入力ファイル名_hanja_YYYYMMDD_HHMM.pdf
```

### `new/convert_hanja_safe.py`

安全版または旧派生版です。

`convert_hanja.py` より前の実装に近く、数字 run のフォントサイズ変更などの処理が含まれています。

### `new/convert_hanja_claude.py`

派生版です。

括弧内文字を 70% 幅にする処理など、現行版とは異なる調整ロジックが含まれています。

### `new/convert_hanjaorg.py`

初期版に近い実装です。

現在の運用では `convert_hanja.py` を優先してください。

## 実行方法

### ドラッグ＆ドロップで実行

1. 変換したい `.docx` ファイルを用意する
2. `new/実行_変換Date.bat` にドラッグ＆ドロップする
3. `new/output` フォルダに変換後の `.docx` と `.pdf` が出力される

### コマンドラインで実行

```bat
python new\convert_hanja.py "input.docx"
```

出力先を指定する場合:

```bat
python new\convert_hanja.py "input.docx" "output.docx"
```

辞書ファイルを明示する場合:

```bat
python new\convert_hanja.py "input.docx" "output.docx" "new\ConvertAllText_improved.xlsm"
```

## 必要な Python パッケージ

最低限必要なパッケージ:

```bat
pip install python-docx lxml openpyxl
```

PDF 出力を有効にする場合:

```bat
pip install docx2pdf
```

`docx2pdf` は Microsoft Word を利用して PDF 変換します。Word が使えない環境では LibreOffice がインストールされていれば代替利用されます。

## 変換処理の流れ

1. 辞書ファイルを探す
   - 明示指定がなければ、スクリプトと同じフォルダ内の `.xlsx` または `.xlsm` を探す
   - ファイル名に `ConvertAllText` を含むものを優先する
2. Excel から辞書を読み込む
   - `Sheet1` から韓国語-漢字辞書を読み込む
   - `Sheet2` から旧字体-新字体表を読み込む
3. Word 文書を読み込む
   - 本文を処理する
   - ヘッダーとフッターも処理する
4. 段落単位で置換する
   - まず単一段落内の語を置換する
   - 次に段落をまたいで分割された語を補正する
   - 最後に旧字体を新字体へ置換する
5. フォントを `Meiryo UI` に統一する
6. 括弧内の文字幅を 95% に圧縮する
7. `.docx` を保存する
8. PDF 変換を試みる

## 置換ロジック

通常の文字列置換では、先に置換した漢字が別の辞書項目に再変換される可能性があります。

そのため `replace_text()` では一度プレースホルダーへ置換し、最後にプレースホルダーを漢字へ戻します。これにより連鎖変換を防いでいます。

辞書は長い語から優先的に置換するため、短い語が先に置換されて長い語の変換を妨げるケースを減らしています。

## 段落またぎ変換

Word 文書では、語の途中で段落が分かれることがあります。

`process_paragraphs()` では、隣接する段落を一時的に連結して、辞書語が段落境界で分割されていないかを確認します。

例:

```text
段落1: 하늘부
段落2: 모님
```

辞書に `하늘부모님` がある場合、境界をまたいだ語として検出し、段落1と段落2に分けて漢字を反映します。

## フォント処理

`set_font_all()` は Word XML 内の run 単位で `w:rFonts` を設定し、以下すべてを `Meiryo UI` にします。

- `ascii`
- `eastAsia`
- `hAnsi`
- `cs`

現行版では段落プロパティ側の `pPr/rPr` は変更せず、段落の配置や均等割付への影響を抑える設計になっています。

## 括弧内文字幅の調整

`condense_parentheses_width()` は、以下の範囲を検出して文字幅を 95% にします。

- 半角括弧: `(...)`
- 全角括弧: `（...）`

括弧内だけに調整をかけるため、必要に応じて Word の run を分割します。

## PDF 出力

PDF 変換は以下の順で試行されます。

1. `docx2pdf`
2. LibreOffice / `soffice`

どちらも使えない場合、`.docx` の保存は完了し、PDF 出力のみスキップされます。

`実行_変換Date.bat` では PDF を必須扱いにしており、最大 120 秒待って PDF が作られない場合はエラーログを作成します。

## 注意点

### 置換前後で文字数が異なる辞書項目

辞書には、置換前後で文字数が異なる項目があります。

例:

```text
선(善) -> 善
악(惡) -> 惡
공생공영공의주의 -> 共生共榮共義
```

`apply_to_segments()` は元の run の文字数を基準に変換後文字列を再配置します。そのため、Word 文書内で細かく run が分かれている箇所では、置換後の文字が別 run の書式を引き継ぐ可能性があります。

### bat の Python パス

bat ファイルでは Python のパスが以下に固定されています。

```bat
%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe
```

Microsoft Store 版 Python を使っていない環境では、このパスが存在しない可能性があります。その場合は bat 内の `PYTHON` 変数を実際の Python パスに変更してください。

### `.xlsm` だがマクロなし

`ConvertAllText_improved.xlsm` は `.xlsm` 形式ですが、現状 VBA は含まれていません。

辞書更新は Excel 上で `Sheet1` と `Sheet2` を編集する運用です。

## 保守時の確認コマンド

Python の構文チェック:

```bat
python -m py_compile new\convert_hanja.py
```

辞書シート構成の確認:

```bat
python -c "import openpyxl; p=r'new\ConvertAllText_improved.xlsm'; wb=openpyxl.load_workbook(p, read_only=True, data_only=True, keep_vba=True); print([(ws.title, ws.max_row, ws.max_column) for ws in wb.worksheets]); wb.close()"
```

## 推奨運用

- 通常利用では `new/実行_変換Date.bat` を使う
- 変換ロジックの修正は `new/convert_hanja.py` に集約する
- 旧版ファイルは比較・退避用として残し、運用対象は明確に `convert_hanja.py` とする
- 辞書を更新した場合は、代表的な Word 文書で変換結果と PDF 出力を目視確認する

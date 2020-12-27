# これはサンプルの Python スクリプトです。

# ⌃R を押して実行するか、ご自身のコードに置き換えてください。
# ダブル⇧ を押すと、クラス/ファイル/ツールウィンドウ/アクション/設定を検索します。

import openpyxl

# ワークブックを新規に作成する
book = openpyxl.Workbook()

# シートを取得し、名前を変更する
sheet = book.active
sheet.title = 'First Sheet'

# 範囲を指定してセルを取得する
cells = sheet['A1': 'B3']
i = 0
for row in cells:
    for cell in row:
        cell.value = i # セルに値を設定する
        i += 1

# ワークブックに名前を付けて保存する
book.save('demo.xlsx')
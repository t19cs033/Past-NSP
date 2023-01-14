import openpyxl as op
from openpyxl.styles import PatternFill
import argparse

# コマンドラインから引数の受け取り，ブックの作成と読み込み
psr  = argparse.ArgumentParser(description='第一引数にASPで導出したファイル，第二引数に設定用ファイルを指定してください')
psr.add_argument('arg1')
psr.add_argument('arg2')
args = psr.parse_args()

book = op.load_workbook(args.arg1)
setting_book = op.load_workbook(args.arg2)

# 設定用ファイルの希望シフトシートを取得
sbs = setting_book.worksheets[3]

# シートを取得しコピー，新たなシートを作成
ws1 = book.copy_worksheet(book.worksheets[0])
# 設定用ファイルに点数の列がないため，合わせるために削除
ws1.delete_cols(5)

ws2 = book.copy_worksheet(book.worksheets[0])

# 名前の変更
ws1.title = '希望シフトの確認'
ws2.title = 'NとJの確認'

# すべてを色なしに
for row in ws1:
  for cell in row:
      cell.fill = PatternFill(fill_type = None)
      ws1[cell.coordinate].font = op.styles.fonts.Font(color='000000')

# 希望シフトがある場合は色を塗る
for row in sbs['L3':'AM34']:
  for cell in row:
      if cell.value != None:
        ws1[cell.coordinate].fill = PatternFill(fgColor='ADFF2F',bgColor="ADFF2F", fill_type = "solid")

for row in ws1:
    if str(row)%2 == 1: 
        insert_num = row + 1
        ws1.insert_rows(insert_num)  
    #if row%2 == 0: 
   #     for cell in row:
    #        ws.cell(row=i, column=j).value = cell.value
'''
#貼り付け開始行
i = 12

#貼り付け開始列
j = 1

#配列ループ
for list in mylist:
    #行をループ
    for row in ws.iter_rows():
        #セルをループ
        for cell in row:
            #配列の数値と同じ行番号だったら
            if cell.row == list:
                #行をコピー
                ws.cell(row=i, column=j).value = cell.value
                j = j + 1
    j = 1
    i = i + 1
'''

# すべてを色なしに
for row in ws2:
  for cell in row:
      cell.fill = PatternFill(fill_type = None)
      ws2[cell.coordinate].font = op.styles.fonts.Font(color='000000')

# NまたはJの場合は色を塗る
for row in ws2['M3':'AN34']:
    for cell in row:
        if cell.value == 'Ｎ':
            cell.fill = PatternFill(fgColor='FF1493',bgColor="FF1493", fill_type = "solid")
        if cell.value == 'Ｊ':
            cell.fill = PatternFill(fgColor='7FFFD4',bgColor="7FFFD4", fill_type = "solid")

# 保存する
book.save('new-4n.xlsx')
# -*- coding: utf-8 -*-
import openpyxl as op
from openpyxl.styles import PatternFill
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl import utils
from copy import copy
import argparse
import datetime

dist_sheet_num = 36
# コマンドラインから引数の受け取り，ブックの作成と読み込み
psr  = argparse.ArgumentParser(description='')
psr.add_argument('arg1')
psr.add_argument('arg2')
args = psr.parse_args()

past_data_book = op.load_workbook(args.arg1)
setting_book = op.load_workbook(args.arg2)
#new_book = op.Workbook()
pdb_1 = past_data_book.worksheets[dist_sheet_num-1] #2022-05-01(35)
pdb_2 = past_data_book.worksheets[dist_sheet_num] #2022-05-29(36)
pdb_3 = past_data_book.worksheets[dist_sheet_num+1] #2022-06-26(37)
sb = setting_book.worksheets[3] #希望シフトがあるシート

#置換する文字列を指定
Henkan_mae = ['期間：',' ～ ４週間','年','月','日']
Henkan_go = ['','','/','/','']
i = 0

for list in Henkan_mae:
    i = i + 1
    #置換
    new_text = pdb_2.cell(row=1,column=1).value.replace(list, Henkan_go[i-1])
    pdb_2.cell(row=1,column=1).value = new_text

sb.cell(row=1,column=2).value = datetime.date(2022, 5, 29)

def PrintShifts(sheet, startcol, endcol, diffcol):
    for i in range(5, 89):
        if i%2 == 1:
            for j in range(startcol, endcol):
                cell_string = str(sheet.cell(row=i, column=j).value)
                if cell_string == "○":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "○"
                if cell_string == "日◇":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "日"
                if cell_string == "J2S":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "S"
                if cell_string == "SS":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "S"
                if cell_string == "N2K":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "N"
                if cell_string == "N2":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "N"
                if cell_string == "J2":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "J"
                if cell_string == "Ｓ":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "S"
                if cell_string == "年":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "年"
                if cell_string == "健":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "健"
                if cell_string == "★":  
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "★"
                if cell_string == "☆":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "☆"
                if cell_string == "P4":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "P"
                if cell_string == "◎":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "◎"
                if cell_string == "":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "／"
                if cell_string == "□":
                    sb.cell(row=(i+1)/2, column=j+diffcol).value = "特"

# shift:転記元のシート名
# startcol:転記元の何列目から読み取るか
# diffcol:転記先の何列目から記載するか

PrintShifts(pdb_1,27,34,-22) #27列目を5列目に記載したいため，diffcol = 5-27 = -22
PrintShifts(pdb_2,6,34,6) #6列目を11列目に記載したいため，diffcol = 11-6+1 = 6
PrintShifts(pdb_3,6,13,34) #6列目を40列目に記載したいため，diffcol = 40-6 = 34

""" # 各シフトの下に希望シフトを追記
for i in range(3, 35):
    ws1.row_dimensions[i*2-2].height = 13 #高さ13に揃える
    for j in range(6, 48):
        ws1.cell(row=i*2-2, column=j).value = sbs.cell(row=i, column=j).value
        cell_string = str(ws1.cell(row=i*2-2, column=j).value)
        ws1.cell(row=i*2-2, column=j).alignment = Alignment(horizontal = 'center', vertical = 'center')
        if cell_string == "S":
            ws1.cell(row=i*2-2, column=j).value = "Ｓ"
        if cell_string == "J":
            ws1.cell(row=i*2-2, column=j).value = "Ｊ"
        if cell_string == "N":
            ws1.cell(row=i*2-2, column=j).value = "Ｎ"
        if cell_string == "P":
            ws1.cell(row=i*2-2, column=j).value = "Ｐ"
        if cell_string == "/":
            ws1.cell(row=i*2-2, column=j).value = "／"
        if j == 6 or j == 13 or j == 41 or j == 48:
            ws1.cell(row=i*2-3, column=j).border = border_ws1_thick_left
            ws1.cell(row=i*2-2, column=j).border = border_dashed_thick_left
        else:
            ws1.cell(row=i*2-3, column=j).border = border_ws1
            ws1.cell(row=i*2-2, column=j).border = border_dashed
        #if ws1.cell(row=i*2-3, column=j).value != None:
            #ws1.cell(row=i*2-3, column=j).fill = PatternFill(fgColor='FF1493',bgColor="FF1493", fill_type = "solid")
        if ws1.cell(row=i*2-2, column=j).value != None and 12 < j and j < 41:
            ws1.cell(row=i*2-3, column=j).fill = PatternFill(fgColor='79c06e',bgColor="FF1493", fill_type = "solid")
            if ws1.cell(row=i*2-2, column=j).value != ws1.cell(row=i*2-3, column=j).value:
                ws1.cell(row=i*2-3, column=j).fill = PatternFill(fgColor='ea553a',bgColor="FF1493", fill_type = "solid")
                ws1.cell(row=i*2-2, column=j).fill = PatternFill(fgColor='ea553a',bgColor="FF1493", fill_type = "solid")
        #ws1.cell(row=i*2-2, column=j).value = op.styles.FORMAT_TEXT 
        #cell_value = ws1.cell(row=i*2-2, column=j).value
        #cell_value.translate(str.maketrans({chr(0x0021 + i): chr(0xFF01 + i) for i in range(94)}))
        #ws1.cell(row=i*2-2, column=j).value = cell_value
        #if ws1.cell(row=i*2-3, column=j).value != sbs.cell(row=i, column=j).value and sbs.cell(row=i, column=j).value != None:
            #ws1.cell(row=i*2-2, column=j).fill = PatternFill(fgColor='FF1493',bgColor="FF1493", fill_type = "solid")
 """
# 無駄な値をクリア
for row in sb.iter_rows(min_row=35, min_col=5, max_row=50, max_col=46):
  for cell in row:
      cell.value = None

str_text = str(new_text)
replace_text = str_text.replace('/','-')
file_name = replace_text + ".xlsx"
# 保存する
setting_book.save(file_name)
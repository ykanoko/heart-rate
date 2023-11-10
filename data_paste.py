import pandas as pd
from openpyxl import load_workbook
import os

for i in range(1,19):
    subject = str(i)
    print('subject', subject)
    # subject = 'S' + str(i + 16)

    for i in range(6):
        type1 = i + 1
        print('type1', type1)

        # CSVファイルの読み込み
        input_dir = r"C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\作業場\\data\\"
        input_file_name = '2309S' + subject + '_' + str(type1) +'.csv.SegTable'
        input_file = input_dir + input_file_name

        if os.path.exists(input_file):
            df_csv = pd.read_csv(input_file, encoding='shift-jis', nrows=12)
            # print("df_csv",df_csv)

            # Excelファイルの読み込み
            # output_dir = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\解析結果\\'
            output_dir = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\解析結果\\未\\'

            output_file_name = '2309S'+ subject +'.xlsx'
            output_file = output_dir + output_file_name
            wb = load_workbook(output_file)

            # 4番目のシートを取得
            sheet_name = wb.sheetnames[type1-1] # シート名を取得
            sheet = wb[sheet_name] # シートを取得

            def write_list_2d(sheet, l_2d, start_row, start_col):
                for y, row in enumerate(l_2d):
                    for x, cell in enumerate(row):
                        sheet.cell(row=start_row + y,
                                column=start_col + x,
                                value=l_2d[y][x])

            write_list_2d(sheet, df_csv.values.tolist(), 2, 1)

            wb.save(output_file)

            print('input_file : ', input_file_name, 'output_file : ', output_file_name, 'sheet_name : ', sheet_name)
        else:
            print('not exist : ', input_file_name)


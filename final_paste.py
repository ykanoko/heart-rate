
from openpyxl import load_workbook
import os

sheet_name_in = 'table'
sheet_name_out = 'Sheet1'


# Excelファイルの読み込み
# 解析結果ディレクトリ
# input_dir = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\解析結果\\'
# 解析結果\結果ディレクトリ
input_dir = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\解析結果\\結果\\'
# 解析結果\未ディレクトリ
# input_dir = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\解析結果\\未\\'

# Excelファイルへの書き込み
output_dir = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\解析結果\\'
output_file_name = 'HR2309.xlsx'
output_file = output_dir + output_file_name


for i in range(19):
    subject = i
    row = i + 3

    # Excelファイルの読み込み
    input_file_name = '2309S'+ str(subject) +'.xlsx'
    input_file = input_dir + input_file_name

    if os.path.exists(input_file):
        print('subject : ', i, 'row', row, 'input_file : ', input_file_name)
        wb_in = load_workbook(filename=input_file, read_only=True, data_only=True)
        sheet_in = wb_in[sheet_name_in]
        print('sheet_in', list(sheet_in.values))

        data = [cell.value for cell in sheet_in['B4':'EO4'][0]]
        print('data', data)

        # Excelファイルへの書き込み
        wb_out = load_workbook(output_file)
        sheet_out = wb_out[sheet_name_out]

        for j, value in enumerate(data, start=2):
            sheet_out.cell(row, column=j, value=value)

        wb_out.save(output_file)




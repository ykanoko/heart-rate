# ディレクトリ内にあるファイル名(例：2309S1_1.csv)をリストに格納し、同じディレクトリ内に新たにtextファイル（ファイル名はファイル作成時の日付+時間）を作成し、そのファイル内にリストを書き込み、保存する。
# 絶対パス「C:\Users\kanok\Documents\大学院\研究室\修論\予備実験230719~\予備実験本番230927~\心拍解析\心拍切り取りデータ\済」

import os
import json
from datetime import datetime

input_dir_name = ['済', '表示方法_済']


dir_path = r'C:\\Users\\kanok\\Documents\\大学院\\研究室\\修論\\予備実験230719~\\予備実験本番230927~\\心拍解析\\心拍切り取りデータ\\'

for dir_name in input_dir_name:
    input_dir = dir_path + dir_name + '\\'

    file_names = os.listdir(input_dir)
    print('dir_name', dir_name, 'file_names', file_names)

    now = datetime.now()
    output_file_name = now.strftime('%Y%m%d_%H%M_') + dir_name + '.txt'

    with open(os.path.join(dir_path, output_file_name), 'w') as f:
        json.dump(file_names, f)

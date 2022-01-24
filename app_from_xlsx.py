import pandas as pd
from glob import glob
import re
import csv


target_xmlx_names = glob(f'target/*.xlsx')
pattern = r'\'【SQL】tbl_datarelation\(構造\)\''
repatter = re.compile(pattern)
#input file name
input_file_name = target_xmlx_names[0]
#xls book Open (xls, xlsxのどちらでも可能)
input_book = pd.ExcelFile(input_file_name)
#sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
input_sheet_name = input_book.sheet_names
#print (input_sheet_name)
datas = dict()


for target_sheet_name in input_sheet_name:
    result = repatter.match(repr(target_sheet_name))
    #もしpatternにmatchするなら読み込み開始
    if result:
        #dict型で格納
        df_sheet_multi = pd.read_excel(input_file_name, sheet_name=[target_sheet_name])
        index = 0 
        #dataの行数を求める
        lens = len(df_sheet_multi[target_sheet_name])
        for len in range(lens):
            data = dict()
            data['name'] = df_sheet_multi[target_sheet_name]['comment'][index]
            data['columnno'] = df_sheet_multi[target_sheet_name]['columnno'][index]
            data['slope'] = df_sheet_multi[target_sheet_name]['slope'][index]
            data['intercept'] = df_sheet_multi[target_sheet_name]['intercept'][index]
            datas[data['columnno']] = data
            index += 1
        #ch順にソート
        get_datas = sorted(datas.items())

with open('result.csv', 'w') as f:
    writer = csv.DictWriter(f, ['name', 'columnno', 'slope', 'intercept'])
    writer.writeheader()
    for get_data in get_datas:
        writer.writerow(get_data[1])
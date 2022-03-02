import csv
from glob import glob
import openpyxl
import pandas as pd
import traceback
import re

import frame

def main(target_xlsm):
    # 初期設定 ####################################################
    # 列指定offset
    row_offset = 12

    # データのあるシートを指定
    pattern_1 = r'\'data_set\''
    repatter_1 = re.compile(pattern_1)
    
    # データのカラム名を指定
    column_1 = 'item'
    column_2 = 'site_id'
    column_3 = 'site_password'
    column_4 = 'old_file_offset'
    column_5 = 'new_file_offset'
    column_6 = 'new_len'

    # Exception_text
    not_found_target_file = '指定のexcelファイルが存在しません'
    fail_dataframe_exchange = 'データフレームへの変換に失敗しました。\nExcelを確認してください'
    
    not_found_write_file = '書き込む.xlsmファイルがありません'
    write_file_two_or_more = '書き込み対象の.xlsmファイルが2個以上存在します'
    unexpected_error = '予期せぬエラーです。管理者にご相談ください。'
    #############################################################

    #input file name
    target_xlsm_name = target_xlsm #glob(f'sample/*.xlsm')[0]
    try:
        #xls book Open (xls, xlsxのどちらでも可能)
        input_file = pd.ExcelFile(target_xlsm_name)
    except:
        frame.show_error(not_found_target_file)
        return False
    #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
    input_file_name =  input_file.sheet_names
    try:
        #DataFrame
        keyword_df = input_file.parse(input_file_name[2],index_col=0)
        keyword_df.fillna('--')
    except:
        frame.show_error(fail_dataframe_exchange)
        return False
    #dict
    keyword_dict = keyword_df.to_dict()
    # keyword 取り出し
    site_id         =  keyword_dict[column_2][column_1]
    site_password   =  keyword_dict[column_3][column_1]
    old_file_offset =  keyword_dict[column_4][column_1]
    new_file_offset =  keyword_dict[column_5][column_1]

    keywords = [ site_id,site_password,old_file_offset,new_file_offset ]
    try:
        #DataFrame
        dataset_df = input_file.parse(input_file_name[1],index_col=0)
        dataset_df.fillna('--')
    except:
        frame.show_error(fail_dataframe_exchange)
        return False
    #dict
    dataset_dict = dataset_df.to_dict()
    # list型でデータ取り出し
    datas = list()
    new_lens        =  dataset_df.loc[:,column_6].values.tolist()
    count_i = 1
    for count_i in range(len(new_lens)):
        count_i += 1
        datas           += [dataset_df.loc[count_i].values.tolist()]
        try:
            #旧サイトの番号をintに変換
            datas[count_i-1][0] = int(datas[count_i-1][0])
        except:
            datas[count_i-1][0] == "--"
        try:
            #新サイトの番号をintに変換
            datas[count_i-1][3] = int(datas[count_i-1][3])
        except:
            datas[count_i-1][3] = '--'
        try:
            #入れ替え用の番号をintに変換
            datas[count_i-1][6] = int(datas[count_i-1][6])
        except:
            datas[count_i-1][6] = '--'
    all_data = [keywords,datas]
 
    return all_data
if __name__=="__main__":
    target_xlsm = glob(f'sample/*.xlsm')[0]
    main(target_xlsm)
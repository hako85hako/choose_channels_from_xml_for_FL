from glob import glob
import csv
import traceback
import pandas as pd
import re
import openpyxl

def main():
    # 初期設定 ####################################################
    target_xlsm = glob(f'sample/*.xlsm')[0]
    
    # 列指定offset
    row_offset = 12

    # データのあるシートを指定
    pattern_1 = r'\'Sheet1\''
    repatter_1 = re.compile(pattern_1)
    # データのカラム名を指定
    
    column_1 = 'comment'
    column_2 = 'columnno'
    column_3 = 'slope'
    column_4 = 'intercept'

    # Exception_text
    not_found_target_file = 'targetファイル内に.xlsxがありません'
    not_correct_target_file = '正しい.xlsxファイルを指定してください'
    not_found_write_file = '書き込む.xlsmファイルがありません'
    write_file_two_or_more = '書き込み対象の.xlsmファイルが2個以上存在します'
    unexpected_error = '予期せぬエラーです。管理者にご相談ください。'
    #############################################################



    #pandasを読み込む
    import pandas as pd

    #input file name
    target_xlsm_name = glob(f'sample/*.xlsm')[0]
    #xls book Open (xls, xlsxのどちらでも可能)
    input_file = pd.ExcelFile(target_xlsm_name)
    #sheet_namesメソッドでExcelブック内の各シートの名前をリストで取得できる
    input_file_name =  input_file.sheet_names
    #lenでシートの総数を確認
    num_sheet = len(input_file_name)
    #DataFrame
    keyword_df = input_file.parse('key_data',index_col=0)
    #dict
    keyword_dict = keyword_df.to_dict()
    # keyword 取り出し
    site_id         =  keyword_dict['site_id']['item']
    site_password   =  keyword_dict['site_password']['item']
    old_file_offset =  keyword_dict['old_file_offset']['item']
    new_file_offset =  keyword_dict['new_file_offset']['item']

    #DataFrame
    dataset_df = input_file.parse('data_set',index_col=0)
    #dict
    dataset_dict = dataset_df.to_dict()
    # list型でデータ取り出し
    datas = list()
    old_lens        =  dataset_df.loc[:,'old_len'].values.tolist()
    count_i = 1
    for count_i in range(len(old_lens)):
        print(count_i)
        count_i += 1
        datas           += [dataset_df.loc[count_i].values.tolist()]
        
    print(datas)#データセットを取得するとこまでできてる



    # old_lens        =  dataset_df.loc[:,'old_len'].values.tolist()
    # old_slopes      =  dataset_df.loc[:,'old_slope'].values.tolist()
    # old_intercepts  =  dataset_df.loc[:,'old_intercept'].values.tolist()
    
    # new_lens        =  dataset_df.loc[:,'new_len'].values.tolist()
    # new_slopes      =  dataset_df.loc[:,'new_slope'].values.tolist()
    # new_intercepts  =  dataset_df.loc[:,'new_intercept'].values.tolist()

    # old_slopes    =  dataset_df.loc[:,'old_slope'].values.tolist()
    # new_lens      =  dataset_df.loc[:,'new_len'].values.tolist()
    # new_slopes    =  dataset_df.loc[:,'new_slope'].values.tolist()





if __name__=="__main__":
   main()
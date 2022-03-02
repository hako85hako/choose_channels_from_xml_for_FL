import csv
import frame
from glob import glob
import re
import os
import shutil
import traceback


import frame

def main(dir_name):
    #初期設定
    not_found_csv = '指定されたフォルダ内にCSVファイルが存在しません'
    file_name_error = 'ファイル名を「FLxxx-xxxxx_0000_yyyy-mm-dd」の形式に統一してください'
    file_name_error_1 = 'ファイル名の先頭は「FL」で統一してください'
    file_name_error_2 = 'FL-IDは「FLxxx-xxxxx」で統一してください'
    file_name_error_3 = 'FL-IDが異なるファイルが含まれています'
    file_name_error_4 = '計測機器IDをFL-IDと日付の間に「_0000_」の形で挿入してください'
    file_name_error_5 = '日付は「yyyy-mm-dd」の形式に統一してください'
    exist_data_file = '指定されたフォルダに、すでにdataフォルダが存在します'
    move_folder_error = 'フォルダ移動中にエラーが発生しました'

    aline_done = 'ファイルの振り分けが完了しました'

    #data以下のファイル格納場所
    target_files = glob(f'{dir_name}/*.csv')
    
    #指定されたフォルダにcsvがない場合
    if not target_files:
        return frame.show_error(not_found_csv)
    
    #ファイル命名規則確認
    try:
        #IDチェック用
        buff_FLID = ""

        for target_file in target_files:
            #ファイル名取り出し
            check_file_name = os.path.split(target_file)[1]
            if not check_file_name[0:2] == 'FL':
                return frame.show_error(file_name_error_1)
            if not re.match('FL[0-9]{3}-[0-9]{5}',check_file_name[0:11]):
                return frame.show_error(file_name_error_2)

            if buff_FLID == "":
                #初回ID格納
                buff_FLID = check_file_name[0:11]
            elif not buff_FLID == check_file_name[0:11]:
                #IDが異なる場合はError
                return frame.show_error(file_name_error_3)

            #buffのIDと今回のIDが一致する場合、計測機器IDチェックへ
            if not check_machine_ID(check_file_name[11:17]):
                return frame.show_error(file_name_error_4)

            #日時チェック
            if not check_date(check_file_name[17:27]):
                return frame.show_error(file_name_error_5)
    except:
        ex = traceback.format_exc()
        return frame.show_error( file_name_error + '\n' + ex )
       
    #dataファルダ作成
    try:
        os.makedirs(dir_name+'/data')
    except:
        return frame.show_error(exist_data_file)
    data_folder = dir_name+'/data'

    try:
        #csvファイルの移動開始
        for target_file in target_files:
            #ファイル名取り出し
            target_file_name = os.path.split(target_file)[1]
            #移動先フォルダ名作成
            create_folder_name = data_folder + '/' +  target_file_name[17:21] + target_file_name[22:24]
            #移動先フォルダ存在確認
            if not os.path.isdir( create_folder_name ):    
                #移動先フォルダ作成
                os.makedirs( create_folder_name )
            #フォルダ移動
            shutil.move( target_file , create_folder_name )
    except:
        ex = traceback.format_exc() 
        return  frame.show_error( move_folder_error + '\n' + ex )
    return frame.show_info(aline_done)
    

def check_machine_ID(machine_ID):
    if re.match('_[0-9]{4}_',machine_ID):
        return True
    else :
        return False

def check_date(date_inf):
    if re.match('[0-9]{4}-[0-9]{2}-[0-9]{2}',date_inf):
        return True
    else :
        return False

# if __name__=="__main__":
#     #テスト用
#     files = ('C:/Users/sakai/Desktop/Project/choose_channel/source/sample/test')
#     main(files)
    
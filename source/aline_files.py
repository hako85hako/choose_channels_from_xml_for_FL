import csv
import frame
from glob import glob
import os
import shutil


import frame

def main(dir_name):
    #data以下のファイル格納場所
    target_files = glob(f'{dir_name}/*.csv')
    
    #指定されたフォルダにcsvがない場合
    if not target_files:
        return frame.show_error('指定されたフォルダ内にCSVファイルが存在しません')
    
    #dataファルダ作成
    os.makedirs(dir_name+'/data',exist_ok=True)
    data_folder = dir_name+'/data'

    #ファイル命名規則確認
    try:
        print('ここに命名規則チェック')
    except:
        return frame.show_error('ファイル名を「FLxxx-xxxxx_0000_yyyy-mm-dd」の形式に統一してください')
    
    #csvファイルの移動開始
    for target_file in target_files:
        #ファイル名取り出し
        target_file_name = os.path.split(target_file)[1]
        #移動先フォルダ名作成
        create_folder_name = data_folder + '/' +  target_file_name[17:21] + target_file_name[22:24]
        #移動先フォルダ存在確認
        if not os.path.isdir( create_folder_name ):    
            try:
                #移動先フォルダ作成
                os.makedirs( create_folder_name )
            except:
                return frame.show_error('予期せぬエラー:0000')

        #フォルダ移動
        shutil.move( target_file , create_folder_name )
    return frame.show_info('ファイルの振り分けが完了しました')
    
if __name__=="__main__":
    #テスト用
    files = ('C:/Users/sakai/Desktop/Project/choose_channel/source/sample/test')
    main(files)
    
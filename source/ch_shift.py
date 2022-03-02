import os
import csv
import frame
from glob import glob

import frame

def main(dir_name,all_data,none_select_column,none_valiable_column):
    #初期設定
    faile_ch_exchange = 'チャンネル置換中にエラーが発生しました'
    not_read_file = 'CSVファイルを読み込めません'
    offset_write_error = 'offset部分の書き込みでエラーが発生しました'
    faile_write_file = 'CSV書き込み中にエラーが発生しました'
    correct_replace =  '完了しました。\n内容を確認してください。' 
    #data以下のファイル格納場所
    files = glob(f'{dir_name}/data/*')
    #各種値格納
    old_file_offset = all_data[0][2] 
    new_file_offset = all_data[0][3]
    password = all_data[0][1]

    for file in files:
        data = []
        #ファイル名作成用########################################################
        FL_ID = all_data[0][0]
        dirname = os.path.basename(file)
        dirname_year = dirname[0:4]
        dirname_month = dirname[4:6]
        dirname = FL_ID + '_' + dirname_year + '-' + dirname_month + '-01'
        #以下形式で出力される
        #FL999-99999_0000_2021-01-01
        #####################################################################
        
         #編集用パス指定
        csv_paths = glob(f'{file}/*.csv')
        for csv_path in csv_paths:
            try:
                #csv読み込み
                f = open(csv_path, 'r')
            except:
                return frame.show_error(not_read_file)
            #headerの取得
            header = next(f)
            #リスト形式
            f = csv.reader(f, delimiter=",", doublequote=True, lineterminator="\r\n", quotechar='"', skipinitialspace=True)
            #listの個数分forを回す
            for row in f:
                #結果の入子
                items = []
                try:
                    for i in range(int(old_file_offset)):
                        #offset8→7の場合
                        if old_file_offset > new_file_offset:
                            if i == 8 :
                                break
                        item = row[i]
                        items += [item]
                        #offset7→8の場合
                        if old_file_offset < new_file_offset:
                            if i == 7 :
                                item = ""
                                items += [item]
                except:
                    return frame.show_error(offset_write_error)

                for ch in all_data[1]:
                    #各種値格納
                    #係数に関しては、該当なければslop=1,intercept=0を格納
                    old_columnno = ch[0]
                    old_slop = 1 if ch[1] == '--' else ch[1]
                    old_intercept = 0 if ch[2] == '--' else ch[2]
                    new_columnno = ch[3]
                    new_slop = 1 if ch[4] == '--' else ch[4]
                    new_intercept = 0 if ch[5] == '--' else ch[5]
                    shift_num = ch[6]

                    if new_columnno == '--':
                        break

                    try:
                        #チャンネルの指定があるかの判定
                        if type(shift_num) is int :
                            #チャンネルの指定あり
                            #指定したチャンネルの中身の判定
                            try:
                                shift_num = shift_num + old_file_offset -1
                                item = float(row[shift_num])
                                #係数の処理
                                item = ( item / old_slop ) - old_intercept
                                item = round(item,7)#DC対応のため、小数点以下7桁まで取得
                                #整数であれば小数点以下を削除
                                if(item.is_integer()):
                                    item = int(item)
                            except:
                                #元の値がnullの場合はそのまま格納する
                                if row[shift_num] == "null" or row[shift_num] == "Null" or row[shift_num] == "\(null\)":
                                    item = row[shift_num]
                                else:
                                    item = none_valiable_column
                        else:
                            item = none_select_column
                    except:
                        return frame.show_error(faile_ch_exchange)
                        item = none_valiable_column
                    items += [item]
                data += [items]

        try:
            data.insert(0, [FL_ID,password])
            path = dir_name+'/'+dirname+'.csv'
            f = open(path, 'w',newline="")
            writer = csv.writer(f)
            writer.writerows(data)
        except:
            return frame.show_error(faile_write_file)
        finally:
            f.close()
        
    return frame.show_info(correct_replace)


# test用
# if __name__=="__main__":
#     target_xlsm = glob(f'sample/*.xlsm')[0]
#     all_data = get_select_ch.main(target_xlsm)
#     dir_name = 'C:/Users/sakai/Desktop/Project/choose_channel/source/sample/test'
#     print(dir_name)
#     main(dir_name,all_data,0,0)
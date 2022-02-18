import os
import csv
import frame
from glob import glob


def main(dir_name,ch_list,none_select_column,none_valiable_column,offset,id,password):
    #data以下のファイル格納場所
    files = glob(f'{dir_name}/data/*')
    #print(files)
    #offset設定
    offset = int(offset)
    for file in files:
        data = []
        #ファイル名作成用########################################################
        dirname = os.path.basename(file)
        dirname_year = dirname[0:4]
        dirname_month = dirname[4:6]
        dirname = id + '_' + dirname_year + '-' + dirname_month + '-01'
        #以下形式で出力される
        #FL999-99999_0000_2021-01-01
        #####################################################################

        #編集用パス指定
        csv_paths = glob(f'{file}/*.csv')
        for csv_path in csv_paths:
            #csv読み込み
            f = open(csv_path, 'r')
            #headerの取得
            header = next(f)
            #リスト形式
            f = csv.reader(f, delimiter=",", doublequote=True, lineterminator="\r\n", quotechar='"', skipinitialspace=True)
            #取得したCSVの列数分forを回す
            for row in f:
                #結果の入子
                items = []
                for i in range(int(offset)+1):
                    item = row[i]
                    items += [item]

                for ch in ch_list:
                    try:
                        #チャンネルの指定があるかの判定
                        if type(ch[0]) is int :
                            #チャンネルの指定あり
                            #指定したチャンネルの中身の判定
                            try:
                                item = float(row[ch[0]])
                                #係数の処理
                                item = ( item * ch[1] ) + ch[2]
                                #item = round(item,4)
                                item = round(item,7)#DC対応のため、小数点以下7桁まで取得
                                #整数であれば小数点以下を削除
                                if(item.is_integer()):
                                    item = int(item)
                            except:
                                #元の値がnullの場合はそのまま格納する
                                if row[ch[0]] == "null" or row[ch[0]] == "Null" or row[ch[0]] == "\(null\)":
                                    item = row[ch[0]]
                                else:
                                    item = none_valiable_column
                        else:
                            item = none_select_column
                    except:
                        raise ValueError("入れ替え時error") 
                    items += [item]
                data += [items]

        data.insert(0, [id,password])
        path = dir_name+'/'+dirname+'.csv'
        f = open(path, 'w',newline="")
        writer = csv.writer(f)
        writer.writerows(data)
        f.close()

def test():
    print('test')


    
if __name__=="__main__":
    frame.TkinterClass()
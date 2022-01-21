import xml.etree.ElementTree as ET 
from glob import glob

def app():
    target_xml_names = glob(f'target/*.xml')
    for target_xml_name in target_xml_names:
        # XMLファイルを解析
        tree = ET.parse(target_xml_name) 
        # XMLを取得
        root = tree.getroot()
        #FL_idを取得
        fl_id = root[0].text
        #passwordを取得
        password = ""
        for pitype in root.iter('Password'):    
            password = pitype.text
        #idとpassを書き込み
        f = open('result/'+fl_id+'_channel_list.txt', 'w',encoding='utf-8', newline='\n')
        f.write(fl_id+','+password+'\n')

        for pitype in root.iter('Channel'):
            if pitype[2].text == '1':
                f.write(pitype[1].text+'\n')

if __name__=="__main__":
    app()
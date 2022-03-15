import app
import csv
import os
import sys
from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import traceback


import aline_files
import app_from_xlsx
import ch_shift
import get_select_ch


class TkinterClass:
    def __init__(self):
        #簡易手順の中身
        
        m1 = '① 本番用FDSManagerから出力した「.xlsx」ファイルを旧サイト、新サイト共に指定する'
        m2 = '② 以下のPathより「ch_list.xlsm」をローカルにコピーし、書き込み用「.xlsm」に指定する'
        m3 = '/fl1/share/部署別フォルダー/技術・生産管理/社員（協力会社）/酒井/酒井 共有/自作ツール/データ移行チャンネル入れ替え'
        m4 = '③ 「Excel出力」を押下すると、新旧サイトのデータが「ch_list.xlsm」に出力される'
        m5 = '④ 「ch_list.xlsm」を開き、W列の「置換」に新サイトチャンネルと紐づけたい変更前チャンネル一覧の「No」を入力する'
        m6 = '  （何も入れない場所は空白）'
        m7 = '⑤ 「置換する直上のフォルダ～」に入れ替え前のCSVが格納されたフォルダを指定する'
        m8 = '  （ファイル名は「FLxxx-xxxxx_○○○○_yyyyy_mm_dd」に統一すること）'
        m9 = '⑥ 最後に残り3項目を選択し、置換開始を押下する事で置換が行われる'

        # ルートを作成
        root = Tk()
        # ''設定
        root.title('Channel_replace')
        root.resizable(True, True)

        # フレーム作成
        frame1 = ttk.Frame(root, padding=(32))
        frame1.grid()

        # label3 = ttk.Label(frame1, text='書き込み用「.xlsm」', padding=(5, 2))
        # label3.grid(row=6, column=0, sticky=E)
############################################################################################
##旧サイト指定用
############################################################################################
        old_xlsx_frame = ttk.Frame(frame1, padding=(5, 10))
        
        old_xlsx_frame_label = ttk.Labelframe(
            old_xlsx_frame,
            text='旧サイト指定用「.xlsx」',
            padding=(10),
            style='My.TLabelframe')
        
        button = ttk.Button(old_xlsx_frame_label, text='参照')
        button.bind('<ButtonPress>', self.old_file_dialog)  
        self.old_file = tk.StringVar()
        self.old_file.set('未選択です')
        label2 = ttk.Label(old_xlsx_frame_label, textvariable=self.old_file, foreground="blue")

        old_xlsx_frame.grid(row=0, column=0, sticky=W)
        old_xlsx_frame_label.grid(row=0, column=0, sticky=E)
        button.grid(row=1, column=0,sticky=W)
        label2.grid(row=1, column=1,sticky=W)
############################################################################################
##新サイト指定用
############################################################################################
        new_xlsx_frame = ttk.Frame(frame1, padding=(5, 10))

        new_xlsx_frame_label = ttk.Labelframe(
            new_xlsx_frame,
            text='新サイト指定用「.xlsx」',
            padding=(10),
            style='My.TLabelframe')
        
        button = ttk.Button(new_xlsx_frame_label, text='参照')
        button.bind('<ButtonPress>', self.new_file_dialog)
        
        self.new_file = tk.StringVar()
        self.new_file.set('未選択です')
        label3 = ttk.Label(new_xlsx_frame_label,textvariable=self.new_file, foreground="blue")
        
        #Layout
        new_xlsx_frame.grid(row=1,column=0,sticky=W)
        new_xlsx_frame_label.grid(row=0, column=0, sticky=E)
        button.grid(row=1, column=0,sticky=W)
        label3.grid(row=1, column=1,sticky=W)
       
############################################################################################
##書き込み用
############################################################################################
        # Frame
        xlsm_frame = ttk.Frame(frame1, padding=(5, 10))

        label2 = ttk.Labelframe(
            xlsm_frame,
            text='書き込み用「.xlsm」',
            padding=(10),
            style='My.TLabelframe')

        
        button = ttk.Button(label2, text='参照')
        button.bind('<ButtonPress>', self.edit_file_dialog)
        self.target_file = tk.StringVar()
        self.target_file.set('未選択です')

        label3 = ttk.Label(label2,textvariable=self.target_file, foreground="blue")


        xlsm_frame.grid(row=2,column=0,sticky=W)
        label2.grid(row=0, column=0, sticky=E)
        button.grid(row=1, column=0,sticky=W)
        label3.grid(row=1, column=1,sticky=W)
############################################################################################
##置換する「data」ファイル
############################################################################################
        # Frame
        data_folder_frame = ttk.Frame(frame1, padding=(5, 10))
        

        label4 = ttk.Labelframe(
            data_folder_frame,
            text='置換するファイル直上の「フォルダ」（この下に「data」フォルダができる）',
            padding=(10),
            style='My.TLabelframe')
        
        
        
        button = ttk.Button(label4, text='参照')
        button.bind('<ButtonPress>', self.data_folder_dialog)
        
        self.data_folder = tk.StringVar()
        self.data_folder.set('未選択です')
        label5 = ttk.Label(label4,textvariable=self.data_folder, foreground="blue")


        data_folder_frame.grid(row=3,column=0,sticky=W)
        label4.grid(row=0, column=0, sticky=E)
        button.grid(row=1, column=0,sticky=W)
        label5.grid(row=1, column=1,sticky=W)

############################################################################################
##取得範囲外のデータ形式
############################################################################################
        # Frame
        oprionFrame2 = ttk.Frame(frame1, padding=(5, 10))
        # Style - Theme
        #ttk.Style().theme_use('classic')
        # Label Frame
        label_frame2 = ttk.Labelframe(
            oprionFrame2,
            text='取得範囲外のデータ形式',
            padding=(10),
            style='My.TLabelframe')

        # Radiobutton 1
        self.v2 = StringVar()
        rb3 = ttk.Radiobutton(
            label_frame2,
            text='null',
            value="null",
            variable=self.v2)

        # Radiobutton 2
        rb4 = ttk.Radiobutton(
            label_frame2,
            text='0',
            value=0,
            variable=self.v2)
        
        # Radiobutton 2
        rb5 = ttk.Radiobutton(
            label_frame2,
            text='空白',
            value='',
            variable=self.v2)

        # Layout
        oprionFrame2.grid(row=5,column=0,sticky=W)
        label_frame2.grid(row=0, column=0)
        rb3.grid(row=0, column=0) # LabelFrame
        rb4.grid(row=0, column=1) # LabelFrame  
        rb5.grid(row=0, column=2) # LabelFrame        


############################################################################################
##取得範囲が空白だった場合に挿入するデータ
############################################################################################
        # Frame
        oprionFrame3 = ttk.Frame(frame1, padding=(5, 10))
        # Style - Theme
        #ttk.Style().theme_use('classic')
        # Label Frame
        label_frame3 = ttk.Labelframe(
            oprionFrame3,
            text='取得先が空白だった場合に挿入するデータ',
            padding=(10),
            style='My.TLabelframe')

        # Radiobutton 1
        self.v3 = StringVar()
        rb6 = ttk.Radiobutton(
            label_frame3,
            text='null',
            value='null',
            variable=self.v3)

        # Radiobutton 2
        rb7 = ttk.Radiobutton(
            label_frame3,
            text='0',
            value=0,
            variable=self.v3)
        
        # Radiobutton 2
        rb8 = ttk.Radiobutton(
            label_frame3,
            text='空白',
            value='',
            variable=self.v3)
    
        # Layout
        oprionFrame3.grid(row=6,column=0,sticky=W)
        label_frame3.grid(row=0, column=0)
        rb6.grid(row=0, column=0) # LabelFrame
        rb7.grid(row=0, column=1) # LabelFrame  
        rb8.grid(row=0, column=2) # LabelFrame        

        endform = ttk.Frame(frame1, padding=(0, 5))
        endform.grid(column=1, sticky=W)

############################################################################################
##係数を有効or無効
############################################################################################
        # Frame
        oprionFrame4 = ttk.Frame(frame1, padding=(5, 10))
        # Style - Theme
        #ttk.Style().theme_use('classic')
        # Label Frame
        label_frame4 = ttk.Labelframe(
            oprionFrame4,
            text='係数の有効化（mame2→mame2であればOFF）',
            padding=(10),
            style='My.TLabelframe')

        # Radiobutton 1
        self.v4 = StringVar()
        rb9 = ttk.Radiobutton(
            label_frame4,
            text='ON',
            value=True,
            variable=self.v4)

        # Radiobutton 2
        rb10 = ttk.Radiobutton(
            label_frame4,
            text='OFF',
            value=False,
            variable=self.v4)
    
        # Layout
        oprionFrame4.grid(row=8,column=0,sticky=W)
        label_frame4.grid(row=0, column=0)
        rb9.grid(row=0, column=0) # LabelFrame
        rb10.grid(row=0, column=1) # LabelFrame  

        endform = ttk.Frame(frame1, padding=(0, 5))
        endform.grid(column=1, sticky=W)

############################################################################################
##簡易手順
############################################################################################
        # Frame
        manual = ttk.Frame(frame1, padding=(5, 10))
        

        manual_tital = ttk.Labelframe(
            manual,
            text='簡易手順',
            padding=(10),
            style='My.TLabelframe')
        
        manual1 = ttk.Label(
            manual_tital,
            text=m1)
        manual2 = ttk.Label(
            manual_tital,
            text=m2)
        
        manual3 = ttk.Entry(manual_tital,width=100)
        manual3.insert(0,m3)

        manual4 = ttk.Label(
            manual_tital,
            text=m4)
        manual5 = ttk.Label(
            manual_tital,
            text=m5)
        manual6 = ttk.Label(
            manual_tital,
            text=m6)
        manual7 = ttk.Label(
            manual_tital,
            text=m7)
        manual8 = ttk.Label(
            manual_tital,
            text=m8)
        manual9 = ttk.Label(
            manual_tital,
            text=m9)

        manual.grid(row=10,column=0,sticky=W)
        manual_tital.grid(row=0, column=0, sticky=E)
        manual1.grid(row=1, column=0,sticky=W)
        manual2.grid(row=2, column=0,sticky=W)
        manual3.grid(row=3, column=0,sticky=W)
        manual4.grid(row=4, column=0,sticky=W)
        manual5.grid(row=5, column=0,sticky=W)
        manual6.grid(row=6, column=0,sticky=W)
        manual7.grid(row=7, column=0,sticky=W)
        manual8.grid(row=8, column=0,sticky=W)
        manual9.grid(row=9, column=0,sticky=W)
############################################################################################
##　ボタン
############################################################################################
        endform = ttk.Frame(manual, padding=(0, 5))
        endform.grid(column=0, sticky=E)
        button1 = ttk.Button(endform, text='Excel出力')
        button1.bind('<ButtonPress>',self.createNewExcel)
        button1.pack(side=LEFT)

        button2 = ttk.Button(endform, text='ファイル振り分け')
        button2.bind('<ButtonPress>',self.move_files)
        button2.pack(side=LEFT)
        
        
        button3 = ttk.Button(endform, text='置換開始')
        button3.bind('<ButtonPress>',self.try_ch_shift)
        button3.pack(side=LEFT)


    
        button3 = ttk.Button(endform, text='閉じる', command=sys.exit)
        button3.pack(side=LEFT)
############################################################################################

        root.mainloop()
    
    def old_file_dialog(self, event):
        fTyp = [("", "*.xlsx")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        old_file = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
        if len(old_file) == 0:
            self.old_file.set('選択をキャンセルしました')
        else:
            self.old_file.set(old_file)

    def new_file_dialog(self, event):
        fTyp = [("", "*.xlsx")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        new_file = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
        if len(new_file) == 0:
            self.new_file.set('選択をキャンセルしました')
        else:
            self.new_file.set(new_file)

    def edit_file_dialog(self, event):
        fTyp = [("", "*.xlsm")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        target_file = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
        if len(target_file) == 0:
            self.target_file.set('選択をキャンセルしました')
        else:
            self.target_file.set(target_file)

    def data_folder_dialog(self, event):
        iDir = os.path.abspath(os.path.dirname(__file__))
        data_folder = tk.filedialog.askdirectory(initialdir=iDir)
        if len(data_folder) == 0:
            self.data_folder.set('選択をキャンセルしました')
        else:
            self.data_folder.set(data_folder)

    #shift_ch指定用のExcel作成
    def createNewExcel(self, event):
        #旧サイトxlsx
        old_xlsx = self.old_file.get()
        #新サイトxlsx
        new_xlsx = self.new_file.get()
        #書き込み用xlsm
        target_xlsm = self.target_file.get()
        try:
            if app_from_xlsx.main(old_xlsx,new_xlsx,target_xlsm) :
                show_info('完了しました。\n内容を確認してください。')
        except:
            ex = traceback.format_exc()
            show_error('処理中にエラーが発生しました。\n\n'+ex)
    
    #ファイル振り分け処理
    def move_files(self,event):
        try:
            #csvの位置を指定
            data_folder = self.data_folder.get()
            #振り分け実行
            aline_files.main(data_folder)
        except:
            ex = traceback.format_exc()
            show_error('振り分け処理中にエラーが発生しました\n\n'+ex)


    #置換処理
    def try_ch_shift(self, event):
        # 書き込んだxlsm
        target_xlsm = self.target_file.get()
        # dataフォルダの位置
        data_folder = self.data_folder.get()
        # 取得範囲外のデータ形式
        none_select_column =  self.v2.get()
        # 入れ替え時対象の数値が空の場合に格納する値の指定
        none_valiable_column = self.v3.get()
        # 係数の有効化
        use_coefficient_flg = self.v4.get()

        try:
            all_data = get_select_ch.main(target_xlsm)
        except:
            ex = traceback.format_exc()
            show_error('Excel読み出し処理中にエラーが発生しました。\n\n'+ex)
        if all_data:
            try:
                ch_shift.main(data_folder,all_data,none_select_column,none_valiable_column,use_coefficient_flg)
            except:
                ex = traceback.format_exc()
                show_error('置換処理中にエラーが発生しました。\n\n'+ex)

    
def show_error(message):
    messagebox.showerror('エラー', message)

def show_info(message):
    messagebox.showinfo('完了',message)

if __name__=="__main__":
   TkinterClass()


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
        label1 = ttk.Label(frame1, text='旧サイト指定用「.xlsx」', padding=(5, 2))
        label1.grid(row=0, column=0, sticky=E)
        button = ttk.Button(frame1, text='参照')
        button.bind('<ButtonPress>', self.old_file_dialog)
        button.grid(row=1, column=1,sticky=E)
        
        self.old_file = tk.StringVar()
        self.old_file.set('未選択です')
        label2 = ttk.Label(frame1,textvariable=self.old_file,width=100)
        label2.grid(row=1, column=2,sticky=E)
############################################################################################
##新サイト指定用
############################################################################################
        label2 = ttk.Label(frame1, text='新サイト指定用「.xlsx」', padding=(5, 2))
        label2.grid(row=3, column=0, sticky=E)
        button = ttk.Button(frame1, text='参照')
        button.bind('<ButtonPress>', self.new_file_dialog)
        button.grid(row=4, column=1,sticky=E)
        
        self.new_file = tk.StringVar()
        self.new_file.set('未選択です')
        label3 = ttk.Label(frame1,textvariable=self.new_file,width=100)
        label3.grid(row=4, column=2,sticky=E)
############################################################################################
##書き込み用
############################################################################################
        label2 = ttk.Label(frame1, text='書き込み用「.xlsm」', padding=(5, 2))
        label2.grid(row=6, column=0, sticky=E)
        button = ttk.Button(frame1, text='参照')
        button.bind('<ButtonPress>', self.edit_file_dialog)
        button.grid(row=7, column=1,sticky=E)
        
        self.target_file = tk.StringVar()
        self.target_file.set('未選択です')
        label3 = ttk.Label(frame1,textvariable=self.target_file,width=100)
        label3.grid(row=7, column=2,sticky=E)
############################################################################################
##置換する「data」ファイル
############################################################################################
        label4 = ttk.Label(frame1, text='置換するファイル直上の「フォルダ」\n（この下に「data」フォルダができる）', padding=(5, 2))
        label4.grid(row=9, column=0, sticky=E)
        button = ttk.Button(frame1, text='参照')
        button.bind('<ButtonPress>', self.data_folder_dialog)
        button.grid(row=10, column=1,sticky=E)
        
        self.data_folder = tk.StringVar()
        self.data_folder.set('未選択です')
        label5 = ttk.Label(frame1,textvariable=self.data_folder,width=100)
        label5.grid(row=10, column=2,sticky=E)

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
        oprionFrame2.grid(row=11,column=1)
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
        oprionFrame3.grid(row=11,column=0)
        label_frame3.grid(row=0, column=0)
        rb6.grid(row=0, column=0) # LabelFrame
        rb7.grid(row=0, column=1) # LabelFrame  
        rb8.grid(row=0, column=2) # LabelFrame        

        endform = ttk.Frame(frame1, padding=(0, 5))
        endform.grid(column=1, sticky=W)
############################################################################################
##　ボタン
############################################################################################
        endform = ttk.Frame(frame1, padding=(0, 5))
        endform.grid(column=6, sticky=W)
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
        error_flg = False
        ###################################################
        # 初期値の設定 #
        ###################################################
        #pathの指定
        #dataディレクトリの場所
        old_xlsx = self.old_file.get()
        new_xlsx = self.new_file.get()
        target_xlsm = self.target_file.get()
        ###################################################

        try:
            if not error_flg:
                if app_from_xlsx.main(old_xlsx,new_xlsx,target_xlsm) :
                    #app.main(dir_name,ch_list,none_select_column,none_valiable_column,offset,id,password)
                    messagebox.showinfo('完了', '完了しました。\n内容を確認してください。')
                else:
                    messagebox.showerror('エラー', 'Excel作成処理中にエラーが発生しました。')
        except:
            ex = traceback.format_exc()
            messagebox.showerror('エラー', '処理中にエラーが発生しました。\n\n'+ex)
    
    #ファイル振り分け処理
    def move_files(self,event):
        #csvの位置を指定
        data_folder = self.data_folder.get()
        #振り分け実行
        aline_files.main(data_folder)


    #置換処理
    def try_ch_shift(self, event):
        error_flg = False
        ###################################################
        # 初期値の設定 #
        ###################################################
        #pathの指定
        #dataディレクトリの場所
        target_xlsm = self.target_file.get()
        data_folder = self.data_folder.get()

        #取得範囲外のデータ形式
        none_select_column =  self.v2.get()
        
        #入れ替え時対象の数値がnull or 空の場合に格納する値の指定
        none_valiable_column = self.v3.get()
        ###################################################
        try:
            all_data = get_select_ch.main(target_xlsm)
        except:
            ex = traceback.format_exc()
            messagebox.showerror('エラー', 'Excel読み出し処理中にエラーが発生しました。\n\n'+ex)

        if all_data:
            try:
                ch_shift.main(data_folder,all_data,none_select_column,none_valiable_column)
                messagebox.showinfo('完了', '完了しました。\n内容を確認してください。')
            except:
                ex = traceback.format_exc()
                messagebox.showerror('エラー', '置換処理中にエラーが発生しました。\n\n'+ex)

    #def aline_files(self,event):

    
def show_error(message):
    messagebox.showerror('エラー', message)

def show_info(message):
    messagebox.showinfo('完了',message)

if __name__=="__main__":
   TkinterClass()


import os
from tkinter import *
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import app
from tkinter import messagebox
import csv
import sys
import traceback

class TkinterClass:
    def __init__(self):
        # ルートを作成
        root = Tk()
        # ''設定
        root.title('csvController')
        root.resizable(True, True)

        # フレーム作成
        frame1 = ttk.Frame(root, padding=(32))
        frame1.grid()

        # ラベル作成
        label1 = ttk.Label(frame1, text=' FL-ID', padding=(5, 2))
        label1.grid(row=0, column=0, sticky=E)

        label2 = ttk.Label(frame1, text='Site-Password', padding=(5, 2))
        label2.grid(row=1, column=0, sticky=E)
############################################################################################
##入れ替え順序指定用CSVの選択
############################################################################################
        label3 = ttk.Label(frame1, text='入れ替え番号指定用CSV', padding=(5, 10))
        label3.grid(row=4, column=0, sticky=E)
        button = ttk.Button(frame1, text='参照')
        button.bind('<ButtonPress>', self.file_dialog)
        button.grid(row=4, column=1,sticky=E)
        
        self.file_name = tk.StringVar()
        self.file_name.set('未選択です')
        label3 = ttk.Label(frame1,textvariable=self.file_name,width=20)
        label3.grid(row=5, column=1,sticky=E)
############################################################################################
##入れ替え順序指定用CSVの選択
############################################################################################
        label4 = ttk.Label(frame1, text='入れ替え番号指定用CSV', padding=(5, 10))
        label4.grid(row=4, column=0, sticky=E)
        button = ttk.Button(frame1, text='参照')
        button.bind('<ButtonPress>', self.file_dialog)
        button.grid(row=4, column=1,sticky=E)
        
        self.file_name = tk.StringVar()
        self.file_name.set('未選択です')
        label4 = ttk.Label(frame1,textvariable=self.file_name,width=20)
        label4.grid(row=5, column=1,sticky=E)
############################################################################################
    
    def file_dialog(self, event):
        fTyp = [("", "*")]
        iDir = os.path.abspath(os.path.dirname(__file__))
        file_name = tk.filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
        if len(file_name) == 0:
            self.file_name.set('選択をキャンセルしました')
        else:
            self.file_name.set(file_name)

    def folder_dialog(self, event):
        iDir = os.path.abspath(os.path.dirname(__file__))
        folder_name = tk.filedialog.askdirectory(initialdir=iDir)
        if len(folder_name) == 0:
            self.folder_name.set('選択をキャンセルしました')
        else:
            self.folder_name.set(folder_name)

    def createNewCSV(self, event):
        error_flg = False
        ###################################################
        # 初期値の設定 #
        ###################################################
        #pathの指定
        #dataディレクトリの場所
        dir_name = self.folder_name.get()
        ###################################################
        
        try:
            if not error_flg:
                print(ch_list)
                app.main(dir_name,ch_list,none_select_column,none_valiable_column,offset,id,password)
                messagebox.showinfo('置換完了', '置換完了しました。\n内容を確認してください。')
        except:
            ex = traceback.format_exc()
            messagebox.showerror('エラー', '置換処理中にエラーが発生しました。\n\n'+ex)
if __name__=="__main__":
   TkinterClass()
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

import app_from_xlsx
import ch_shift


class TkinterClass:
    def __init__(self):
        # ルートを作成
        root = Tk()
        # ''設定
        root.title('choose_channel')
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
##　ボタン
############################################################################################
        endform = ttk.Frame(frame1, padding=(0, 5))
        endform.grid(column=6, sticky=W)
        button1 = ttk.Button(endform, text='Excel出力')
        button1.bind('<ButtonPress>',self.createNewExcel)
        button1.pack(side=LEFT)

        button2 = ttk.Button(endform, text='置換開始')
        button2.bind('<ButtonPress>',ch_shift.test())
        button2.pack(side=LEFT)
    
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
        except:
            ex = traceback.format_exc()
            messagebox.showerror('エラー', '処理中にエラーが発生しました。\n\n'+ex)
    
    def call_except_in_app(ex):
        messagebox.showerror('エラー', '処理中にエラーが発生しました。\n\n'+ex)
if __name__=="__main__":
   TkinterClass()
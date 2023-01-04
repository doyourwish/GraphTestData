# coding: utf-8

from cProfile import label
import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
import openpyxl
from functools import partial

def main():
    print("two_window.py start")
    #------------------------------------------
    #    カラム選択ウィンドウ生成
    #------------------------------------------
    #メインウィンドウ生成
    main_win = tk.Tk()
    main_win.title("テスト用ウィンドウ")
    main_win.geometry("400x300")
    
    #エクセルからカラム取得
    getdata = GetData("data/book.xlsx")
    column = getdata.GetColumns()
    
    #list取得
    #list_varを関数の返り値にすると、型がおかしくなってできない？
    list = getdata.GetList(column)
    list_var = tk.StringVar(value=list)

    #selected_column @listboxで選択されたカラム名を保持、グラフ表示に使用
    #label_list @listboxで表示したラベルを保持、listboxエリアの表示クリアに使用
    selected_column=[]
    label_list=[]
    
    #listbox生成
    #listbox @listvariableで指定した項目を表示する
    #listbox.bind @項目選択時に第二引数の関数を実行
    listbox = tk.Listbox(main_win, height=5, listvariable=list_var)
    listbox.bind('<<ListboxSelect>>', lambda e: ShowLabel(listbox,selected_column,label_list))
    listbox.pack()
    
    #クリアボタン生成
    #tk.Buttonのcommand引数は、()をつけるとエラーが発生する
    #lambda式で書くとエラーは吐かない
    button = tk.Button(text='クリア', font=("Meiryo", "12", "normal"), command = lambda : ClearLabelList(selected_column,label_list))
    button.place(x=100, y =250, width=100, height=50 )
    #閉じるボタン生成
    button = tk.Button(text='終了', font=("Meiryo", "12", "normal"), command = lambda : WindowDestroy(main_win))
    button.place(x=100, y =150, width=100, height=50 )
    
    #メインウィンドウ表示
    main_win.mainloop()

    #------------------------------------------
    #    グラフ選択ウィンドウ生成
    #------------------------------------------
    index_tmp=[]
    #ボタンが抑えれた時の関数？
    def check_list():
        check =cb.get()
        index_tmp.append(check)
        print(index_tmp)
    #メイン画面を作成
    window=tk.Tk()
    #サイズを設定
    window.geometry('300x200')
    #タイトル設定
    window.title('Combo box')
    #コンボボックスを設置
    print(selected_column)
    cb=ttk.Combobox(window,values=selected_column)
    cb.pack(pady=50)
    #ボタンを設置
    btn=tk.Button(window,text="決定",command=check_list)
    btn.place(x=140,y=100)
    window.mainloop()
    
#******************************************
#    define class and function
#******************************************
class GetData:
    def __init__(self,path):
        print("[GetData::__init__] constractor start")
        self.path = path
        self.book = openpyxl.load_workbook(self.path)

    def GetColumns(self):
        sheets=self.book.sheetnames
        df=[]
        for i in range(len(sheets)):
            df.append(pd.read_excel(self.path,sheet_name=sheets[i]))
        column=[]
        #TODO:複数シート対応
        column=df[0].columns
        return column
    
    def GetList(self,column):
        list=[]
        for i in range(len(column)):
            list.append(column[i])   
        print(list)
        return list 
    
#リストを選択すると、ラベルに選択した項目が追加される
def ShowLabel(listbox,selected_column,label_list):
    selected_index = listbox.curselection()
    selected_module = listbox.get(selected_index)
    label = tk.Label(text=selected_module)
    label.pack()
    
    print(selected_module)
    selected_column.append(selected_module)
    label_list.append(label)
    
#ボックス内をクリアしてリセットして再度検索可能にする
def ClearLabelList(selected_column,label_list):
    for i in range(len(label_list)):
        label_list[i].pack_forget()
    label_list.clear()
    selected_column.clear()
    print("[ClearLabelList] clear box's list")
    
#ウィンドウを閉じる
def WindowDestroy(window):
    window.destroy()
    print("[WindowDestroy] window destory")

#******************************************
#    main
#******************************************
if __name__ == '__main__':
    main()

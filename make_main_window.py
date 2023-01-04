# coding: utf-8

import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,NavigationToolbar2Tk
import glob
import os
import shutil

def main():
    print("make_main_window.py start")

    #メインウィンドウ生成
    main_win =tk.Tk()
    main_win.title("graph window")
    main_win.geometry("600x800")
    
    #エクセルからカラム取得
    getdata = GetData("data/book.xlsx")
    column = getdata.GetColumns()
    
    #list取得
    #list_varを関数の返り値にすると、型がおかしくなってできない？
    list = getdata.GetList(column)
    list_var = tk.StringVar(value=list)
    
    #初期化
    #selected_column @listboxで選択されたカラム名を保持、グラフ表示に使用
    #label_list @listboxで表示したラベルを保持、listboxエリアの表示クリアに使用
    selected_column=[]
    label_list=[]
    label_list2=[]
    label_list3=[]
    
    #クリアボタン
    button = tk.Button(text='clear', font= ("MS Gothic","10", "normal"), command = lambda : clear(label_list))
    button.place(relx = 0.1,rely = 0.9 ,anchor = tk.CENTER,width = 70, height = 30)
    
    #caseラベル、ボタン設置
    case_label = tk.Label(text="CASE番号を,区切りで入力")
    case_label.place(x=30,y=30)
    textBox4= tk.Entry(width=20)
    textBox4.place(x=30,y=50)
    button = tk.Button(text='clear', font= ("MS Gothic", "10","normal"), command = lambda : reset(textBox4))
    button.place (x=115, y=70,width= 40, height = 20)

    #ラジオボタン設置
    radio_label=tk.Label(text="第二軸も使用")
    radio_label.place(x=30,y=280)
    radio_value=tk.IntVar(value=1)
    radio_yes=tk.Radiobutton(text ="Yes",command = lambda : radio_click(radio_value,label_list3,main_win),variable=radio_value, value=0)
    radio_no=tk.Radiobutton(text = "No" ,command = lambda : radio_click(radio_value,label_list3,main_win),variable=radio_value, value=1)
    radio_yes.place(x=30,y=300)
    radio_no.place(x=30,y=320)

    #グラフ化ボタン
    button = tk.Button(text="graph",font =("MS Gothic","10", "normal"), command = lambda : graph(textBox4,radio_value,label_list))
    button.place(relx = 0.5,rely = 0.9, anchor = tk.CENTER,width = 70, height = 30)
    
    #閉じるボタン
    button = tk.Button(text='end', font= ("MS Gothic","10", "normal"), command = lambda : end(main_win))
    button.place(relx = 0.9,rely = 0.9 ,anchor = tk.CENTER, width = 70, height = 30)

    #listラベル、ボックス設置
    list_label = tk.Label(text="CSV列名")
    list_label.place(x=270,y=10)
    listbox=tk.Listbox(width = 30,height=20, listvariable=list_var)
    listbox.bind('<<ListboxSelect>>', lambda e: select_list(listbox,selected_column,label_list))
    listbox.place(x=200,y=30)

    #unitラベル、リストボックス設置
    unit_label=tk.Label(text="グラフの縦軸の単位選択")
    unit_label.place(x=30,y=160)
    list_unit=["大気温","温度","回転数"]
    list_var2=tk.StringVar(value=list_unit)
    listbox2=tk.Listbox(width=24,height=5, listvariable=list_var2,selectmode="single")
    listbox2.bind("<<ListboxSelect>>", lambda e: select_unit(listbox2,label_list2))
    listbox2.place(x=30,y=180)

    #updateする理由は？
    #これがないとスクロールが出てこないが、よくわかってない
    main_win.update()

    #スクロールバー設置
    scroll2=tk.Scrollbar(orient=tk.VERTICAL,command=listbox2.yview)
    scroll2.place(x=160,y=180,height=listbox2.winfo_height())
    listbox2["yscrollcommand"]=scroll2.set
    scroll=tk.Scrollbar(orient=tk.VERTICAL,command=listbox.yview)
    scroll.place(x=380,y=30,height=listbox.winfo_height())
    listbox['yscrollcommand']= scroll.set

    main_win.mainloop()


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

def clear(label_list):
    label_list.clear()
    for i in range(len(label_list)):
        label_list[i].destroy()
    print("[clear] clear box's list")
        
def end(window):
    window.destroy()
    print("[end] window destory")

def reset(textbox):
    textbox.delete(0,tk.END)

#リスト選択
def select_list(listbox,selected_column,label_list):
    selected_index = listbox.curselection()
    selected_module = listbox.get(selected_index)
    label = tk.Label(text=selected_module)
    label.pack(side=tk.TOP,anchor=tk.E)
    
    print("[select_list] selected_module =",selected_module)
    selected_column.append(selected_module)
    label_list.append(label)
    
    
#ユニット選択
#先頭に配列を追加してる理由は？
def select_unit(listbox,label_list):
    selected_index = listbox.curselection()
    #listboxの項目を選んだ状態で別のlistboxの項目を選択すると、何故か関数が両方のlistboxから呼ばれる
    #選択されてないlistboxから関数を呼び出すことは防げてないが、選択は解除されるのでそれで判別
    if len(selected_index)!=0:
        for i in selected_index:
            print("[select_unit] selected_index[i] =",i)
        selected_module = listbox.get(selected_index)
        label_list.insert(0,selected_module)
        print("[select_unit] label_list =",label_list)

def mean():
   add_column = textBox.get()
   column_num = textBox4.get()
   column_list = []
   listbox.insert(tk.END, add_column)
   while True:
       check = column_num.find(",")
       if check==-1:
           column_list.append(column_num)
           break
       column_list.appned(column_num[check-1])
       column_num = column_num[check+1:]
       for i in range(len(column_list)):
            mean_num = column_list[i]
            mean_num = int(mean_num)-1
            df_tmp[mean_num][add_column]=0
            for j in range(len(selected_column)):
                df_tmp[mean_num][add_column] += df_tmp[mean_num][selected_column[j]]
            df_tmp[mean_num][add_column] = (df_tmp[mean_num][add_column])/len(selected_column)

def radio_click(radio_value,label_list,window):
    radio=radio_value.get()
    if radio==0:
        #ラベル生成
        radio_label1 = tk.Label(text="第1軸でplotする配列番号を、区切りで入力")
        radio_label1.place(x=30,y=355)
        textBox5 = tk.Entry(width=20)
        textBox5.place(x=30,y=380)
        radio_label2 = tk.Label(text="第2軸でplotする配列番号を、区切りで入力")
        radio_label2.place(x=30,y=400)
        textBox6 = tk.Entry(width=20)
        textBox6.place(x=30,y=420)
        unit_label3 = tk.Label(text="グラフの第2軸の単位選択")
        unit_label3.place(x=30,y=450)
        
        first_column = textBox5.get()
        second_column = textBox6.get()
        
        #リストボックス生成
        list_unit3 = ["a","b","c","d","e"]
        list_var3 = tk.StringVar(value=list_unit3)
        listbox3 = tk.Listbox(width = 24,height=5, listvariable=list_var3, selectmode= "single")
        listbox3.bind('<<ListboxSelect>>', lambda e: select_unit(listbox3,label_list))
        listbox3.place(x=30,y=470)
        
        #updateする理由は？
        #これがないとスクロールが出てこないが、よくわかってない
        window.update()
        
        #スクロールバー生成
        scroll3=tk.Scrollbar(orient=tk.VERTICAL, command=listbox3.yview)
        scroll3.place(x=160,y=470,height=listbox3.winfo_height())
        listbox3 ["yscrollcommand"] = scroll3.set
        first_column=textBox5.get()
        second_column=textBox6.get()
        
    #グラフ用に返り値を返す
    return radio,first_column,second_column
        
def check_column(column,list):
    print("[check_column] column =",column)
    while True:
        check=column.find(",")
        if check==-1:
            list.append(column)
            break
        list.append(column[check-1])
        column=column[check+1:]
        
def graph(textBox4,radio_value,label_list):
    #グラフ表示用サブウィンドウ生成
    sub_win = tk.Toplevel()
    sub_win.title("graph_window")
    sub_win.geometry("1000x600")
    
    #case_num取得
    case_num=textBox4.get()
    radio,first_column,second_column=radio_click(radio_value,label_list,sub_win)
    print(radio)
    print(first_column)
    print(second_column)
    
    #カラムチェック
    first_list=[]
    second_list=[]
    case_list=[]
    check_column(first_column,first_list)
    check_column(second_column,second_list)
    check_column(case_num,case_list)
        
    if len(case_list)>1:
        for j in range(len(selected_column)):
            fig = plt.Figure (figsize(12,4))
            ax = fig.add_subplot(111)
            for i in range(len(case_list)):
                number=case_list[i]
                print("[graph,len(case_list)>1] case_list =",case_list[i])
                number=int(number)-1
                ax.plot(df_tmp[number][selected_column[j]], label="Case"+str (number+ 1))
            ax.set_xlabel("経過時間(s)",fontname="MS Gothic")
            ax.set_ylabel(label_list2[0],fontname="MS Gothic")
            ax.grid()
            ax.legend(loc=(1.05, 0.7))
            fig.suptitle(selected_column[j],fontname="MS Gothic")
            fig.subplots_adjust(right =0.8)
            canvas = FigureCanvasTkAgg(fig,sub_win)
            canvas.draw()
            canvas.get_tk_widget().pack()
            fig,path=" ./graphy/"+selected_column[i]+".png"
            if os.path.exists(fig_path):
                name, ext =os.path.splitext(fig_path)
                i=1
                while True:
                    new_path="{}{:0=3}{}".format(name,i,ext)
                    if not os.path.exists(new_path):
                        fig.savefig(new_path)
                        break
                    i+=1
    else:
        if radio==1:
            case_num= int(case_num)-1
            fig = plt.Figure(figsize=(12,4))
            ax = fig.add_subplot(111)
            for i in range(len(selected_column)):
                ax.plot(df_tmp[case_num][selected_column[i]],label=selected_column[i])
            ax.set_xlabel("経過時間" ,fontname="MS Gothic")
            ax.set_ylabel(label_list2[0],fontname= "MS Gothic")
            ax.grid()
            ax.legend(loc= (1.05, 0.7) ,prop={"family": "MS Gothic"},fontsize=10)
            fig.suptitle("Case"+str(case_num+1), fontname= "MS Gothic")
            fig.subplots_adjust(right=0.7)
            canvas = FigureCanvasTkAgg(fig,sub_win)
            canvas.draw()
            canvas.get_tk_widget().pack()
            fig_path="./graph/"+"Case"+str(case_num+1)+" png"
            if os.path.exists(fig_path):
                name,ext =os.path.splitext(fig_path)
                i=1
                while True:
                    new_path="{}{:0=3}{}".format(name,i,ext)
                    if not os.path.exists(new_path):
                        fig.savefig(new_path)
                        break
                    i+=1
            else:
                fig.savefig(fig_path)
        #同一case内で第二動を使ってplot
        else:
            fig=plt.Figure(figsize=(12,4))
            ax1=fig.add_subplot(111)
            ax2=ax1.twinx()
            for i in range(len(first_list)):
                number=first_list[i]
                print("[graph,radio!=1] radio =",radio)
                print("[graph,radio!=1] first_list =",first_list[i])
                number=int(number)-1
                ax1.plot(df_tmp[number][selected_column[number]])
            for i in range(len(second_list)):
                number=second_list[i]
                print("[graph] second_list =",second_list[i])
                number=int(number)-1
                ax2.plot(df_tmp[number][selected_column[number ]])
            ax1.set_xlabel("経過時間",fontname="MS Gothic")
            ax1.set_ylabel(label_list2[0],fontname="MS Gothic")
            ax2.set_xlabel("経過時間",fontname="MS Gothic")
            ax2.set_ylabel(label_list3[0],fontname="MS Gothic")
            ax1.grid()
            ax1.legend (loc=(1.05, 0.7) ,prop= {"family": "MS Gothic"},fontsize=10)
            ax2.grid()
            ax2.lezend(loc= (1.05, 0.8) ,prop={"family": "MS Gothic"},fontsize= 10)
            fig.suptitle ("Case",fontname="MS Gothic")
            fig.subplots_adjust(right=0.7)
            canvas = FigureCanvasTkAgg(fig,sub_win)
            canvas.draw()
            canvas.get_tk_widget().pack()
            fig_path="./graph/"+"Case"+str(case_num+1)+" png"
            if os.path.exists(fig_path):
                name,ext =os.path.splitext(fig_path)
                i=1
                while True:
                    new_path="{}{:0=3}{}".format(name,i,ext)
                    if not os.path.exists(new_path):
                        fig.savefig(new_path)
                        break
                    i+=1
            else:
                fig.savefig(fig_path)
        sub_win.mainloop()

#******************************************
#    main
#******************************************
if __name__ == '__main__':
    main()

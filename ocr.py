import tkinter.messagebox
import os
import cv2
from paddleocr import PaddleOCR
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import *
import re
import xlrd
import matplotlib.pyplot as plt
import numpy as np

# Paddleocr supports Chinese, English, French, German, Korean and Japanese.
# You can set the parameter `lang` as `ch`, `en`, `fr`, `german`, `korean`, `japan`
# to switch the language model in order.


ocr = PaddleOCR(use_angle_cls=True, lang='ch')  # need to run only once to download and load model into memory
count1 = 0
count2 = 0
list=[]
def ocrtest():
    global count0
    global count1
    global count2
    global h
    count1=0
    count2=0
    count0=0
    for filename in os.listdir('D:/23/huang'):
        gro=()
        print(filename)
        img_path = cv2.imread('D:/23/huang/' + filename)
        result = ocr.ocr(img_path, cls=True)
        txts = [line[-1][-2] for line in result]
        str1 = ",".join(txts)
        print(str1)
       # m : str=r'姓名：,(.{2,3}),采样时间：,(\d+-\d+-\d{2} {0,1}\d+:\d+:\d{2}).{40,60}检测结果：,(.{4,7}),'
        m: str = r'姓名：,(.{2,3}),采样时间：,(\d+-\d+-\d{2} {0,1}\d+:\d+:\d{2}).{6,30},检测时间：,(\d+-\d+-\d{2} {0,1}\d+:\d+:\d{2}|待检测机构上传),.{5,16},检测结果：,(.{4,7}),'
        r = re.search(m, str1,re.M|re.I)
        print(r.groups())
        gro+=r.groups()
        count0 += 1
        if r.group(4)=='【阴性】':
            count1 +=1
        if r.group(4)=='待检测机构上传':
            count2 +=1
        #.append(r.groups())
        print(r.group(2))
        print(r.group(3))
        if r.group(3)=='待检测机构上传':
            h='结果未出'
            list5 =[]
            list5.append(h)
            h1=tuple(list5)
            print(h1)
            gro+=(h1)
        else:
            h=0
            res = max(r.groups(), key=len, default='')
            print(len(r.group(2)))
            print(len(r.group(3)))
            if len(r.group(2))==19 and len(r.group(3))==19:
                for i in range(len(res)):
                    if (i == 9) and r.group(2)[i] !=r.group(3)[i]:
                            h += 24 * ((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
                    if (i == 11) and r.group(2)[i] !=r.group(3)[i]:
                            h += 10 *((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
                    if (i == 12) and r.group(2)[i] !=r.group(3)[i]:
                            h += ((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
            if len(r.group(2))==18 and len(r.group(3))==19:
                for i in range(len(res)):
                    if (i == 9) and r.group(2)[i] !=r.group(3)[i]:
                            h += 24 * ((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
                    if (i == 11) and r.group(2)[i-1] !=r.group(3)[i]:
                            h += 10 *((int(r.group(3)[i])) - int(r.group(2)[i-1]))
                            continue
                    if (i == 12) and r.group(2)[i-1] !=r.group(3)[i]:
                            h += ((int(r.group(3)[i])) - int(r.group(2)[i-1]))
                            continue
            if len(r.group(2))==19 and len(r.group(3))==18:
                for i in range(len(res)):
                    if (i == 9) and r.group(2)[i] !=r.group(3)[i]:
                            h += 24 * ((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
                    if (i == 11) and r.group(2)[i] !=r.group(3)[i-1]:
                            h += 10 *((int(r.group(3)[i-1])) - int(r.group(2)[i]))
                            continue
                    if (i == 12) and r.group(2)[i] !=r.group(3)[i-1]:
                            h += ((int(r.group(3)[i-1])) - int(r.group(2)[i]))
                            continue
            if len(r.group(2))==18 and len(r.group(3))==18:
                for i in range(len(res)):
                    if (i == 9) and r.group(2)[i] !=r.group(3)[i]:
                            h += 24 * ((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
                    if (i == 10) and r.group(2)[i] !=r.group(3)[i]:
                            h += 10 *((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
                    if (i == 11) and r.group(2)[i] !=r.group(3)[i]:
                            h += ((int(r.group(3)[i])) - int(r.group(2)[i]))
                            continue
            print(h)
            list4 =[]
            list4.append(str(h))
            h2 = tuple(list4)
            print(h2)
            gro+=(h2)
        list.append(gro)
        print(gro)
    print(list)
    print('阴性结果：'+str(count1)+'  未出结果：'+str(count2))
    return list,count1,count2

def sort():
    list.sort(reverse=True, key=lambda x: x[1])
    print(list)
    return list

def out():
    for i in range(len(list)):
        tv.insert('','end',values=list[i])

def delete():
    tv.delete(tv.selection())

def focus():
    print(tv.focus())

def delButton():
    x=tv.get_children()
    for item in x:
        tv.delete(item)

def pluss():
    tup1 = tuple([etName.get()])+tuple([etTime.get()])+tuple([etJtime.get()])+tuple([etResult.get()])
    print(tup1)
    list.append(tup1)
    print(list)

def dormre():
    x = tv.focus()
    print(x[-1])
    a=int(x[-1])
    data = xlrd.open_workbook('D:/23/dorm.xls')
    table = data.sheet_by_index(0)
    print(list[a-1][0])
    for i in range(9):
        if list[a-1][0] == table.cell(i, 0).value:
            print(table.cell(i, 1).value)
            tk.messagebox.showinfo('结果', table.cell(i, 1).value)

def plot():
    s1 = round(100 * count1 / (count1 + count2))
    s2 = round(100 * count2 / (count1 + count2))
    print(s1,s2)
    y = np.array([s1,s2])
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.pie(y,
            labels=['阴性', '结果未出'],  # 设置饼图标签
            colors=["#d5695d", "#5d8ca8"])
    plt.title("核算结果饼图")
    plt.show()

#print(ocrtest())
#print(sort())



win = tk.Tk()
win.title("核酸结果管理信息系统")
win.geometry('1000x550')
#btOcrtest = tk.Button(win,text="识别图像",command=ocrtest)
#btSort = tk.Button(win,text="排序",command=sort)
#btOcrtest.grid(row=0, column=0, padx=5, pady=5)
#btSort.grid(row=1, column=0, padx=5, pady=5)

area=('姓名','采样时间','检测时间','结果','间隔时间（小时）')
ac=('all','m','c','e','s')
tv=ttk.Treeview(win,columns=ac,show='headings',height=18)
tv.column(ac[0],width=150,anchor='center')
tv.heading(ac[0],text=area[0])
tv.column(ac[1],width=300,anchor='center')
tv.heading(ac[1],text=area[1])
tv.column(ac[2],width=200,anchor='center')
tv.heading(ac[2],text=area[2])
tv.column(ac[3],width=210,anchor='center')
tv.heading(ac[3],text=area[3])
tv.column(ac[4],width=160,anchor='center')
tv.heading(ac[4],text=area[4])
tv.pack()


frame1 = Frame (win, relief=RAISED, borderwidth=2)
frame1 .pack(side=BOTTOM, fill=BOTH, ipadx=13, ipady=13, expand=0)
Button(frame1,text="识别图像",command=ocrtest) .pack(side=LEFT, padx=15, pady=13,expand=YES)
Button(frame1, text="排序",command=sort) .pack(side=LEFT, padx=15, pady=13,expand=YES)
Button(frame1, text="显示",command=out) .pack (side=LEFT, padx=15, pady=13,expand=YES)
Button(frame1, text="删除",command=delete) .pack (side=LEFT, padx=15, pady=13,expand=YES)
Button(frame1, text="清空",command=delButton) .pack (side=LEFT, padx=15, pady=13,expand=YES)
Button(frame1, text="宿舍查询",command=dormre) .pack (side=LEFT, padx=15, pady=13,expand=YES)
Button(frame1, text="画图",command=plot) .pack (side=LEFT, padx=15, pady=13,expand=YES)


frame2 = Frame (win, relief=RAISED, borderwidth=2)
frame2 . pack (side=BOTTOM, fill=X, ipadx="10", ipady="10", expand=1)
etName = tk.StringVar()
etTime = tk.StringVar()
etJtime = tk.StringVar()
etResult = tk.StringVar()

Label (frame2, text="名字：") .pack (side=LEFT, padx="10", pady="15")
Entry (frame2, textvariable=etName) .pack (side=LEFT, padx="10", pady="15")
Label (frame2,text="采样时间：") .pack (side=LEFT, padx="10", pady="15")
Entry (frame2, textvariable=etTime) .pack (side=LEFT, padx="10", pady="15")
Label (frame2,text="检测时间：") .pack (side=LEFT, padx="10", pady="15")
Entry (frame2, textvariable=etJtime) .pack (side=LEFT, padx="10", pady="15")
Label (frame2,text="检测结果：") .pack (side=LEFT, padx="10", pady="15")
Entry (frame2, textvariable=etResult) .pack (side=LEFT, padx="10", pady="15")
Button(frame2, text="加行",command=pluss) .pack (side=LEFT, padx=13, pady=13)

win.mainloop()


dataframe = pd.DataFrame(list)
print(dataframe)
dataframe.to_excel('D:/23/result.xls')
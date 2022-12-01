import pandas as pd
import xlwings as xw
import tkinter.ttk
from tkinter import *
import tkinter.messagebox as msg
import tkinter.filedialog
import sys
root = Tk()
root.title('请选择文件')
root.geometry('400x300')
show_folderPath1 = Entry(root)
show_folderPath1.grid(row = 3,column = 2)
show_folderPath2 = Entry(root)
show_folderPath2.grid(row = 2,column = 2)
show_folderPath3 = Entry(root)
show_folderPath3.grid(row = 1,column = 2)
show_folderPath4 = Entry(root)
show_folderPath4.grid(row = 5,column = 2)

def get_work():
    filename=show_folderPath1.get()
    datalist = xz()
    x = pd.read_excel(filename, sheet_name=None)# 读取需拆分文件路径
    for i in range(len(datalist)):
     app = xw.App(visible=False, add_book=False)  # 启动Excel程序窗口，但不新建工作簿
     workbook = app.books.add()  # 新建一个工作簿。
     workbookname=datalist[i]
     workbook.save(workbookname + '.xlsx')  # 在工作空间文件夹下生成工作簿
     workbook.close()  # 关闭工作簿
     app.quit()  # 退出Excel程序
     app.kill()  # 清除Excel进程
     writer = pd.ExcelWriter(workbookname + '.xlsx', engine='openpyxl')
     for j in x.keys():
      name = str(j)
      sheetData = get_excel_sheet_data(filename, name)  # 读取每个sheet下的数据
      Ename=datalist[i]
      titlename=get_list_title()#获取列标题
      needdata=select_need_data(sheetData,Ename,titlename)#获取符合条件的数据
      wk = pd.DataFrame(needdata)
      wk.to_excel(writer, sheet_name=name)
     writer.close()
    root.destroy()#关闭窗口
    sys.exit();
    msg.showinfo(title='消息提示',message='文件已拆分完毕！')

def xz():
 fzfilename =show_folderPath2.get()
 app = xw.App(visible=False, add_book=False)
 app.display_alerts = False
 app.screen_updating = False
 wb = app.books.open(fzfilename)
 sheet = wb.sheets[0]
 data = sheet.range('A1').expand().value#获取辅助文件数据
 return data
 wb.close()
 app.quit()  # 退出Excel程序
 app.kill()  # 清除Excel进程

def get_list_title():
    title = show_folderPath3.get()#获取输入的列标题数据
    return title;

def get_excel_sheet_data(filename,sheetname):
    try:
        df = pd.read_excel(filename, sheet_name=sheetname,header=0,index_col=None)  #读取sheet数据
        dataList = df.to_dict(orient='records');
        return dataList;
    except:
        msg.showinfo(title='消息提示', message='未能打开此文件，请确认文件名及路径是否正确')
        # root.destroy();
        return []

def select_need_data(basedata,Exname,titlename): #筛选所需数据
    try:
        filt = list(filter(lambda x: x[titlename] == Exname, basedata))#筛选数据，只获取符合条件的数据
        print(filt);
        return filt;
    except:
        msg.showinfo(title='消息提示', message='未读取到数据或文件列标题有误，请检查文件数据！')
        root.destroy();
        sys.exit();
        return []

def fz():#辅助文件路径
    fzFile = tkinter.filedialog.askopenfilename(title="Select Excel file", filetypes=(("Excel files","*.xlsx;*.xls;*.xlsm"),))
    show_folderPath2.delete(0, END)  # 清空
    show_folderPath2.insert(0, fzFile)  # 写入路径
    return fzFile

def dz():#拆分文件路径
    dzFile = tkinter.filedialog.askopenfilename(title="Select Excel file", filetypes=(("Excel files","*.xlsx;*.xls;*.xlsm"),))  # 选择文件路径
    show_folderPath1.delete(0, END)  # 清空
    show_folderPath1.insert(0, dzFile)  # 写入路径
    return dzFile
def hz():#需要合并的文件所在文件夹路径
    hzFile=tkinter.filedialog.askdirectory()
    show_folderPath4.delete(0, END)  # 清空
    show_folderPath4.insert(0, hzFile)  # 写入路径
    return hzFile

def get_select_check():#按需求执行拆分程序
    if(b4_var.get()==1):
     easy_work()  #简单拆分
    elif(b4_var.get()==0):
     get_work()   #复杂条件拆分
    else:
     msg.showinfo(title='消息提示', message='未匹配到执行逻辑，请检查文件数据及操作规范！')

def get_excel_all():
       filename=show_folderPath4.get()


def state_change():#变更输入法及按钮状态
    if(b4_var.get()==1):
     show_folderPath3['state']=DISABLED
     show_folderPath2['state']=DISABLED
     btn2['state'] = DISABLED
    else:
     btn2['state'] = ACTIVE
     show_folderPath3['state'] = NORMAL
     show_folderPath2['state'] = NORMAL
# 勾选不同选项相应组件的状态变更
def cf_change():
    if(b6_var.get()==1):
        btn5['state'] = DISABLED
        btn6['state'] = DISABLED
        btn2['state'] = ACTIVE
        btn1['state'] = ACTIVE
        btn2['state'] = ACTIVE
        btn3['state'] = ACTIVE
        show_folderPath3['state'] = NORMAL
        show_folderPath2['state'] = NORMAL
        show_folderPath1['state'] = NORMAL
        show_folderPath4['state'] = DISABLED
        b7.deselect()
        b4['state'] = ACTIVE

    else:
        btn5['state'] = ACTIVE
        btn2['state'] = ACTIVE
        btn6['state'] = ACTIVE
        show_folderPath4['state'] = DISABLED
        show_folderPath3['state'] = DISABLED
        show_folderPath2['state'] = DISABLED
        show_folderPath1['state'] = DISABLED
        b4['state'] = DISABLED

def hb_change():
    if (b7_var.get() == 1):
        show_folderPath3['state'] = DISABLED
        show_folderPath2['state'] = DISABLED
        show_folderPath1['state'] = DISABLED
        show_folderPath4['state'] = NORMAL
        btn1['state']=DISABLED
        btn2['state'] = DISABLED
        btn3['state'] = DISABLED
        btn5['state'] = ACTIVE
        btn6['state'] = ACTIVE
        b4['state'] = DISABLED
        b6.deselect()
        b4.deselect()
    else:
        btn2['state'] = ACTIVE
        show_folderPath3['state'] = NORMAL
        show_folderPath2['state'] = NORMAL
        show_folderPath1['state'] = NORMAL
        btn1['state'] = ACTIVE
        btn2['state'] = ACTIVE
        btn3['state'] = ACTIVE
        b4['state'] = ACTIVE
# 简单拆分函数
def easy_work():
    filename = show_folderPath1.get()
    if(filename==""):
        msg.showinfo(title='消息提示', message='文件路径不能为空，请重新输入！')
    else:
     x = pd.read_excel(filename, sheet_name=None)  # 读取需拆分文件路径
     for j in x.keys():
      name = str(j)
      writer = pd.ExcelWriter(name + '.xlsx', engine='openpyxl')
      sheetData = get_excel_sheet_data(filename, name)  # 读取每个sheet下的数据
      wk = pd.DataFrame(sheetData)
      wk.to_excel(writer, sheet_name=name)  #写入excel文件
      writer.close();
     msg.showinfo(title='消息提示', message='文件已拆分完毕！')
     root.destroy()  # 关闭窗口
     sys.exit();#close the task



# excel文件拆分模块
b4_var=tkinter.IntVar()
b6_var=tkinter.IntVar()
b7_var=tkinter.IntVar()
w=Label(root,text="请输入筛选列标题：")
w.grid(row=1,column = 1)
b4=Checkbutton(root,text="简单拆分",variable=b4_var,onvalue=1, offvalue=0,command=state_change)
b4.grid(row=1,column=3)

b2=Label(root,text="请选择辅助文件：")
b2.grid(row = 2,column = 1)
btn2=Button(root,text="浏览文件",command=fz,state=NORMAL)
btn2.grid(row = 2,column = 3)

b1=Label(root, text="请选择需拆分文件：")
b1.grid(row = 3,column = 1)
btn1=Button(root, text="浏览文件", command=dz)
btn1.grid(row = 3,column = 3)

btn3=Button(root, text="开始拆分", command=get_select_check)
btn3.grid(row = 4,column = 1)
# excel文件合并模块
btn6=Button(root, text="开始合并(开发中)", command=get_excel_all)
btn6.grid(row = 4,column = 2)

b5=Label(root,text="请选择需合并文件文件夹")
b5.grid(row=5,column=1)
btn5=Button(root,text="浏览文件夹",state=NORMAL,command=hz)
btn5.grid(row=5,column=3)

b6=Checkbutton(root,text="拆分模式",variable=b6_var,onvalue=1, offvalue=0,command=cf_change)
b6.grid(row=0,column=1)

b7=Checkbutton(root,text="合并模式(开发中)",variable=b7_var,onvalue=1, offvalue=0,command=hb_change)
b7.grid(row=0,column=2)
root.mainloop()#保持窗口

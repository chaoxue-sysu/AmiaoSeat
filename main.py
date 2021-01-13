# @Date: 2021/01/12
# @Author: Xue Chao
# @Description:
import tkinter.filedialog
import tkinter.messagebox
from tkinter import *
from tkinter.font import Font
import os
import openpyxl
import random

def new_win():
    tkinter.messagebox.showinfo(title='提示', message='运行成功！点击 “确定” 查看结果！')
    check()

def set_priority(seats,students_num):
    col_row=seats
    cols=sorted(col_row.keys())
    print(f'{len(cols)} col seats')
    print(cols)
    all_seats=0
    for col in cols:
        all_seats+=len(col_row[col])
    priors=[]
    skip=False
    for i in range(min(cols),max(cols)+1):
        if skip or not cols.__contains__(i):
            i+=1
            skip=False
            continue
        priors.append(i)
        skip=True
    if all_seats/2<students_num:
        col_pro={}
        for i in cols:
            if not priors.__contains__(i):
                idx=cols.index(i)
                try:
                    f=cols[idx]-cols[idx-1]
                except:
                    f=2
                try:
                    n=cols[idx+1]-cols[idx]
                except:
                    n=2
                col_pro[i]=n+f
        for x in sorted(col_pro.keys(),key=lambda x:(-col_pro[x],x)):
            priors.append(x)
    return priors

def seat(path,out_path):
    workbook=openpyxl.load_workbook(path)
    sh1=workbook.worksheets[1]
    seat_sh=workbook.worksheets[0]
    names={}
    name_idx=[]
    for i in range(sh1.max_row):
        for j in range(sh1.max_column):
            na=str(sh1.cell(i+1,j+1).value).strip()
            if na.__contains__('姓名'):
                name_idx.append([i,j])
    for r,j in name_idx:
        for i in range(r+1,sh1.max_row):
            try:
                name=sh1.cell(i+1,j+1).value.strip()
            except:
                name=''
            if name=='':
                continue
            names[name]=[i+1,j-2+1]
    nums={}
    for i in range(3,seat_sh.max_row):
        for j in range(seat_sh.max_column):
            try:
                snum=int(seat_sh.cell(i+1,j+1).value)
                nums[snum]=[i+1,j+1]
            except:
                continue
    if len(nums)/2<len(names):
        print(f'warning: {path} not ok!: info: {len(nums)} seats vs {len(names)} students')
    else:
        print(f'info:{path}  {len(nums)} seats vs {len(names)} students')

    col_row={}
    for seat in nums.keys():
        r,c=nums[seat]
        if not col_row.__contains__(c):
            col_row[c]=[]
        col_row[c].append([r,seat])
    priors=set_priority(col_row,len(names))
    print(priors)
    i=-1
    finish=False
    stu_names=sorted(names.keys(),key=lambda x:random.random())
    for col in priors:
        if finish:
            break
        for row,sid in col_row[col]:
            if finish:
                break
            i+=1
            if i>=len(stu_names):
                finish=True
                continue
            seat_sh.cell(row+1,col,stu_names[i])
            sh1.cell(*names[stu_names[i]],sid)
    workbook.save(out_path)


def do():
    out_dir=str(entry2.get()).strip()
    try:
        files=in_files
    except:
        files=[]
    if len(files)==0 or out_dir=='':
        tkinter.messagebox.showerror(title='错误', message='您还没有选择输入文件或者输出路径哦！')
        return
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    for p in files:
        path=p
        print(path)
        if not os.path.exists(path):
            continue
        np='.'.join(os.path.basename(p).split('.')[:-1]+['amiu','xlsx'])
        out_path=f'{out_dir}/{np}'
        seat(path,out_path)
    new_win()

def select_multi():
    files=tkinter.filedialog.askopenfilenames(title='选择文件（可多选）', filetypes=[
        ("Excel format", ".xlsx"),
    ])
    x=';'.join(list(files))
    global in_files
    in_files=list(files)
    entry1.delete(0,END)
    entry1.insert(0,x)

def select_dir():
    dir=tkinter.filedialog.askdirectory()
    entry2.delete(0,END)
    entry2.insert(0,dir)

def check():
    os.startfile(entry2.get())
    # tkinter.filedialog.askopenfilename()

def main():
    global entry1,entry2
    root = Tk()
    # root.geometry('480x280')
    root.title('阿喵排排座 V1.0')
    entry1=Entry(root,width=50,font=Font(family='Times', size=10))
    entry2=Entry(root,width=10,font=Font(family='Times', size=10))
    root.iconbitmap('logo.ico')
    lablet=Label(root,text='阿喵排排座 V1.0',font=Font(family='Times', size=24, weight='bold'),fg='black')
    lable1=Label(root,text='输入.xlsx文件（可多选）',font=Font(family='Times', size=12, weight='bold'),width=25)
    sele_btn1=Button(root,command=select_multi,text='选择文件',font=Font(family='Times', size=10, weight='bold'),bg='pink',fg='black',width=8)

    lable2=Label(root,text='输出目录',font=Font(family='Times', size=12, weight='bold'),width=10)
    sele_btn2=Button(root,command=select_dir,text='选择目录',font=Font(family='Times', size=10, weight='bold'),bg='pink',fg='black',width=8)

    btn=Button(root,command=do,text='点击开始排座',font=Font(family='Times', size=15, weight='bold'),bg='pink',fg='black',width=20)


    lablet.grid(row=0,column=0,pady=35,sticky=N+S+E+W,columnspan=3)
    lable1.grid(row=1,column=0,sticky=N+S+E+W)
    entry1.grid(row=1,column=1,sticky=N+S+E+W)
    sele_btn1.grid(row=1,column=2,sticky=N+S+E+W)
    lable2.grid(row=2,column=0,sticky=N+S+E+W, pady=5)
    entry2.grid(row=2,column=1,sticky=N+S+E+W, pady=5)
    sele_btn2.grid(row=2,column=2,sticky=N+S+E+W, pady=5)
    btn.grid(row=3,column=0,columnspan=3,sticky=N+S,pady=30)

    # curWidth = root.winfo_width()
    # curHight = root.winfo_height()
    curWidth,curHight=640,280
    scn_w, scn_h = root.maxsize()
    cen_x = (scn_w - curWidth) / 2
    cen_y = 0.8*(scn_h - curHight) / 2
    size_xy = '%dx%d+%d+%d' % (curWidth, curHight, cen_x, cen_y)
    root.geometry(size_xy)
    root.mainloop()


if __name__=='__main__':
    main()


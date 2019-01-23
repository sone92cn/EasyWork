#encoding="utf8"
import os, time
from tkinter import E, S, W, N
from PathTool import createTreeAsPath
from tkinter import Tk, ttk, filedialog, messagebox, StringVar, IntVar

predir = ""
def getNumString(num, length):
    s = "%d" % num
    k = length - len(s)
    if k > 0:
        s = ("0" * k) + s
    return s

def renameFiles(infiles, rtype, vlen, ent1, ent2, ent3, rev, rdel, dlen, rcut, clen):
    if len(ent2.get()):
        reverse = rev.get()
        if reverse:
            nth = int(ent2.get()) + len(infiles) - 1
        else:
            nth = int(ent2.get())
        for file in infiles:
            fpath = os.path.split(file)[0]
            fname = os.path.split(file)[1]
            i = fname.rfind(".")
            nstr = getNumString(nth, vlen)
            if rtype == 1:
                outfile = os.path.join(fpath, ent1.get() + nstr + ent3.get() + fname[i:])
            elif rtype == 2:
                outfile = os.path.join(fpath, ent1.get() + nstr + ent3.get() + fname)
            elif rtype == 3:
                outfile = os.path.join(fpath, fname[:i] + ent1.get() + nstr + ent3.get() + fname[i:])
            os.rename(file, outfile)
            if reverse:
                nth -= 1
            else:
                nth += 1
    else:
        for file in infiles:
            fpath = os.path.split(file)[0]
            fname = os.path.split(file)[1]
            i = fname.rfind(".")
            if rtype == 2:
                outfile = os.path.join(fpath, ent1.get() + ent3.get() + fname)
                print(outfile)
            elif rtype == 3:
                outfile = os.path.join(fpath, fname[:i] + ent1.get() + ent3.get() + fname[i:])
            else:
                continue
            os.rename(file, outfile)

def getFilesToRename(rbx, cbx, com1, ent1, ent2, ent3, rev, cbx3, com3, cbx4, com4):
    global predir
    infiles = filedialog.askopenfilenames(title="请选择需要重命名的文件：", initialdir = predir, filetypes=(("Image Files", "*.jpg;*.jpeg;*.png"), ("Excel Files", "*.xls;*.xlsx;*.xlsm"), ("Document Files", "*.doc;*.docx"), ("ALL Files", "*.*")))
    if len(infiles):
        if cbx.get():
            renameFiles(sorted(infiles), rbx.get(), int(com1.get()), ent1, ent2, ent3, rev, rbx3.get(), int(com3.get()), rbx4.get(), int(com4.get()))
        else:
            renameFiles(sorted(infiles), rbx.get(), 0, ent1, ent2, ent3, rev, rbx3.get(), int(com3.get()), rbx4.get(), int(com4.get()))
        predir = os.path.split(infiles[0])[0]
        messagebox.showinfo("提示", "操作已完成")

def getFolderToRename(rbx, cbx, com1, ent1, ent2, ent3, rev, cbx3, com3, cbx4, com4):
    global predir
    indir = filedialog.askdirectory(title="请选择需要重命名的文件所在的目录：", initialdir = predir)
    if len(indir):
        infiles = createTreeAsPath(indir, fileRegular = r".*\.(jpg|jpeg|png|xls|xlsx|xlsm|doc|docx|pdf)$", scanSubFolder = False, treeMode = False)
        if len(infiles):
            if cbx.get():
                renameFiles(sorted(infiles), rbx.get(), int(com1.get()), ent1, ent2, ent3, rev, rbx3.get(), int(com3.get()), rbx4.get(), int(com4.get()))
            else:
                renameFiles(sorted(infiles), rbx.get(), 0, ent1, ent2, ent3, rev, rbx3.get(), int(com3.get()), rbx4.get(), int(com4.get()))
        predir = indir
        messagebox.showinfo("提示", "操作已完成")
   
def main():
    month = getNumString(time.localtime().tm_mon, 2)
    
    root = Tk()
    root.title("文件重命名")
    #left=wx.DisplaySize()[0] - 200
    
    mainframe = ttk.Frame(root, padding="3 3 12 12")
    mainframe.columnconfigure(0, weight=1)
    mainframe.rowconfigure(0, weight=1)
    mainframe.grid(row=0, column=0, sticky=(N, W, E, S))
    
    lab1 = ttk.Label(mainframe, text="前缀：")
    lab2 = ttk.Label(mainframe, text="起始编号：", padding="10 0 0 0")
    lab3 = ttk.Label(mainframe, text="后缀：")
    labe = ttk.Label(mainframe, text="    ")
    
    str1 = StringVar(value="CFW_CQ"+month+"_")
    str2 = StringVar(value="1")
    str3 = StringVar()
    ent1 = ttk.Entry(mainframe, textvariable=str1)
    ent2 = ttk.Entry(mainframe, textvariable=str2)
    ent3 = ttk.Entry(mainframe, textvariable=str3)
    
    rbx = IntVar(value=1)
    rbx1 = ttk.Radiobutton(mainframe, variable=rbx, text='重新命名', value = 1)
    rbx2 = ttk.Radiobutton(mainframe, variable=rbx, text='添加到原文件名前', value = 2)
    rbx3 = ttk.Radiobutton(mainframe, variable=rbx, text='追加到原文件名后', value = 3)
    rbx4 = ttk.Radiobutton(mainframe, variable=rbx, text='原文件名删除：', value = 4)
    rbx5 = ttk.Radiobutton(mainframe, variable=rbx, text='原文件名保留：', value = 5)
    
    cbv1 = IntVar(value=1)
    com1 = ttk.Combobox(mainframe, values = ["2", "3", "4", "5", "6", "7", "8"], width=2, state='readonly')
    cbx1 = ttk.Checkbutton(mainframe, text = "固定编号长度到：", variable = cbv1, onvalue = 1, offvalue = 0)
    labx = ttk.Label(mainframe, text="位")

    rev = IntVar(value=0)
    com2 = ttk.Checkbutton(mainframe, text = "使用倒序方式编号", variable = rev, onvalue = 1, offvalue = 0)

    cbv3 = IntVar(value=0)
    choi = "-5,-4,-3,-2,-1,1,2,3,4,5"
    com3 = ttk.Combobox(mainframe, values = choi.split(","), width=2, state='normal')
    laby = ttk.Label(mainframe, text="位")

    cbv4 = IntVar(value=0)
    com4 = ttk.Combobox(mainframe, values = choi.split(","), width=2, state='normal')
    labz = ttk.Label(mainframe, text="位")
    
    button1= ttk.Button(mainframe, text="选择文件")
    button2= ttk.Button(mainframe, text="选择目录")
    button1.bind('<1>',lambda e: getFilesToRename(rbx, cbv1, com1, ent1, ent2, ent3, rev, cbv3, com3, cbv4, com4))
    button2.bind('<1>',lambda e: getFolderToRename(rbx, cbv1, com1, ent1, ent2, ent3, rev, cbv3, com3, cbv4, com4))
    
    lab1.grid(row=0, column=0, sticky=(N, E, S))
    labe.grid(row=0, column=4, sticky=(N, E, S))
    lab2.grid(row=1, column=0, sticky=(N, E, S))
    lab3.grid(row=2, column=0, sticky=(N, E, S))
    labe.grid(row=0, column=4, sticky=(N, E, S))
    ent1.grid(row=0, column=1, columnspan=3, sticky=(N, W, S))
    ent2.grid(row=1, column=1, columnspan=3, sticky=(N, W, S))
    ent3.grid(row=2, column=1, columnspan=3, sticky=(N, W, S))
    rbx1.grid(row=3, column=1, columnspan=3, sticky=(N, W, S))
    rbx2.grid(row=4, column=1, columnspan=3, sticky=(N, W, S))
    rbx3.grid(row=5, column=1, columnspan=3, sticky=(N, W, S))
    rbx4.grid(row=6, column=1, sticky=(N, W, S))
    com3.grid(row=6, column=2, sticky=(N, W, S))
    laby.grid(row=6, column=3, sticky=(N, E, S))
    rbx5.grid(row=7, column=1, sticky=(N, W, S))
    com4.grid(row=7, column=2, sticky=(N, W, S))
    labz.grid(row=7, column=3, sticky=(N, E, S))
    cbx1.grid(row=8, column=1, sticky=(N, W, S))
    com1.grid(row=8, column=2, sticky=(N, W, S))
    labx.grid(row=8, column=3, sticky=(N, E, S))
    com2.grid(row=9, column=1, sticky=(N, W, S))

    
    
    
    button1.grid(row=10, column=1, sticky=(N, W, S)) 
    button2.grid(row=11, column=1, sticky=(N, W, S)) 
    
    root.geometry("280x290+500+230")
    
    com1.current(0)
    root.mainloop()
            
if __name__ == "__main__":
    main()

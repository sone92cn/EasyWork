#version=2016.10.29.001

from tkinter import Tk, ttk, messagebox, StringVar, DoubleVar, E, S, W, N

class Calculator(object):
    def __init__(self, entbox, entvar, labvar):
        self.index = -1
        self.calcstr = []
        self.entbox = entbox
        self.entvar = entvar
        self.labvar = labvar
        self.vartax6 = False
        self.calcold = ""
        self.v = 0
    def calcprev(self):
        if self.index > 0:
            self.index -= 1
            self.entvar.set(self.calcstr[self.index])
            self.labvar.set(0.0)
            self.entbox.selection_range(0, 0)
            self.entbox.icursor(len(self.entvar.get()))
    def calcnext(self):
        if self.index > -1 and self.index < len(self.calcstr) - 1:
            self.index += 1
            self.entvar.set(self.calcstr[self.index])
            self.labvar.set(0.0)
            self.entbox.selection_range(0, 0)
            self.entbox.icursor(len(self.entvar.get()))
    def calcfunc(self):
        calcstr = "self.v=(" + self.entvar.get() + ")"
        try:
            exec(calcstr)
        except:
            self.labvar.set(0.0)
        else:
            self.calcold = calcstr
            self.labvar.set(round(self.v, 4))
            self.entbox.selection_range(0, len(self.entvar.get()))
            self.calcstr.append(self.entvar.get())
            if len(self.calcstr) > 10:
                self.calcstr.pop(0)
            self.index = len(self.calcstr) - 1
            self.entbox.icursor(len(self.entvar.get()))
            if win32clipboard is not None:
                try:
                    win32clipboard.OpenClipboard()
                except:
                    messagebox.showinfo("提示", "计算结果未拷贝！")
                else:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(str(round(self.v, 2)))
                    win32clipboard.CloseClipboard()
    def calcsele(self):
        if self.entvar.get():
            self.entbox.selection_range(0, len(self.entvar.get()))
            self.entbox.icursor(len(self.entvar.get()))
    def calccopy(self):
        if self.labvar.get():
            self.entvar.set(str(self.labvar.get()))
            self.entbox.selection_range(0, 0)
            self.entbox.icursor(len(self.entvar.get()))
    def calctax6(self):
        if len(self.calcold) < 9 or self.calcold[7:] != self.entvar.get()[:len(self.calcold)-7]:
            self.vartax6 = False
            self.calcold = "self.v=(" + self.entvar.get() + ")"
        if self.vartax6:
            calcstr = self.calcold + "/1.06*0.06"
        else:
            calcstr = self.calcold + "/1.06"
        try:
            exec(calcstr)
        except:
            self.labvar.set(0.0)
        else:
            self.labvar.set(round(self.v, 4))
            self.entvar.set(calcstr[7:])
            self.vartax6 = not self.vartax6
            self.entbox.icursor(len(self.entvar.get()))
            self.entbox.selection_range(0, len(self.entvar.get()))
        if win32clipboard is not None:
                try:
                    win32clipboard.OpenClipboard()
                except:
                    messagebox.showinfo("提示", "计算结果未拷贝！")
                else:
                    win32clipboard.EmptyClipboard()
                    win32clipboard.SetClipboardText(str(round(self.v, 2)))
                    win32clipboard.CloseClipboard()

if __name__ == "__main__":
    root = Tk()
    root.title("计算器")
    root.wm_attributes('-topmost',1)
    root.geometry(root.geometry('0x0-200+100'))
    
    frame = ttk.Frame(root, padding="3 3 12 12")
    frame.grid(row=0, column=0, sticky=(N, W, E, S))
    
    calcnum = DoubleVar()
    calcstr = StringVar()
    calcbox = ttk.Entry(frame, width=25, textvariable=calcstr, font=("黑体", 20, "bold"), justify="right")
    calclab = ttk.Label(frame, width=25, textvariable=calcnum, font=("黑体", 20, "bold"), justify="right", anchor="e")
    calclab.grid(row=0, column=0, sticky=(W, N, E))
    calcbox.grid(row=1, column=0, sticky=(W, N, E))
    
    calc = Calculator(calcbox, calcstr, calcnum)
    calcbox.bind("<Key-Up>", lambda e: calc.calcprev())
    calcbox.bind("<Key-Down>", lambda e: calc.calcnext())
    calcbox.bind("<Key-Return>", lambda e: calc.calcfunc())
    calcbox.bind("<Key-F6>", lambda e: calc.calctax6())
    calcbox.bind("<Key-F11>", lambda e: calc.calcsele())
    calcbox.bind("<Key-F12>", lambda e: calc.calccopy())
	
    try:
        import win32clipboard
    except:
    	messagebox.showinfo("提示", "无法访问粘贴板，计算结果将无法自动复制！")

    calcbox.focus()
    root.mainloop(0)

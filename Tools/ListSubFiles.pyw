"""扫描文件夹的子目录，并保存到CSV文件"""
import csv
from FelixFunc import isNumber
from PathTool import createTreeAsPath
from tkinter import Tk, filedialog, messagebox

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    indir = filedialog.askdirectory(title="请选择需要列出子文件夹的目录：")
    if len(indir):
        infiles = createTreeAsPath(indir, fileRegular="", scanSubFolder=False, treeMode=False, forFile=True, relativePath=True)
        if len(infiles):
            fname = filedialog.asksaveasfilename(title="另存为", filetypes=(("CSV File", "*.csv"), ))
            if len(fname):
                if len(fname) < 4:
                    fname += ".csv"
                elif fname[-4:] != ".csv":
                    fname += ".csv"
                with open(fname, "w", newline="") as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(["Scan In Folder:",""])
                    writer.writerow(["",indir])
                    writer.writerow(["Get Sub Files:",""])
                    for infile in infiles:
                        if '.' in infile:
                            n = infile.rfind('.')
                            flnam = infile[:n]
                            flext = infile[n:]
                        else:
                            flnam = infile
                            flext = ''
                        if isNumber(infile):
                            flnam = "'" + flnam
                            infile = "'" + infile
                        elif isNumber(flnam):
                            flnam = "'" + flnam
                        writer.writerow(["", infile, flnam, flext])
                messagebox.showinfo("提示", "结果已储存到：" + fname)
            else:
                messagebox.showinfo("提示", "无法储存结果，操作取消！")
        else:
            messagebox.showinfo("提示", "选择的文件夹无文件，操作取消！")
    else:
        messagebox.showinfo("提示", "未选择文件夹，操作取消！")
    root.destroy()
 

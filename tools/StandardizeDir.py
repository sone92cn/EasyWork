"""脚本功能：将目录规范化为同一目录下仅能包含文件夹或文件，不能同时都有的形式"""
import os, shutil
from tkinter import Tk, filedialog

def doStandardizeDir(path, scanSubfolder = False):
    flist = []
    hasFile = False
    hasFolder = False
    for file in os.listdir(path):
        filepath = os.path.join(path, file)
        if os.path.isfile(filepath):
            hasFile = True
            flist.append([1, filepath])
        elif os.path.isdir(filepath) and file[:1] != ".":
            hasFolder = True
            flist.append([0, filepath])
    
    failed = []
    move = False
    npath = os.path.join(path, os.path.split(path)[1])
    if hasFile and hasFolder:
        move = True
        if not os.path.exists(npath):
            os.mkdir(npath)
    for file in flist:
        if file[0]:
            if move:  
                nfile = os.path.join(npath, os.path.split(file[1])[1])
                if not os.path.exists(nfile):
                    try:
                        shutil.move(file[1], nfile)
                        print("Succeed to move: " + file[1] + "\n    To:" + nfile)
                    except:
                        failed.append(file[1])
                else:
                    failed.append(file[1])
        elif scanSubfolder:
            lst = doStandardizeDir(file[1], scanSubfolder)
            if len(lst):
                failed.extend(lst)
    return failed

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    indir = filedialog.askdirectory(title="请选择要整理的目录：")
    if len(indir):
        print("整理目录：" + indir)
        y = input("确定整理该文件夹吗（Y/N）：")
        if y == "Y" or y == "y":
            failed = doStandardizeDir(indir, True)
            if len(failed):
                print("\n------------------------------------------")
                for file in failed:
                    print("Fail to move: " + file)
        os.system("pause")
    root.destroy()

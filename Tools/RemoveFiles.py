"""删除目录下，满足regular（正则表达式）指定文件名的文件"""
import os, sys
from tkinter import Tk, filedialog
from PathTool import createTreeAsPath

"""指定需要清理的文件"""
regular = '^.*\.backup$'  #'^~.*$'

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="请选择需要清理的目录：", initialdir=sys.path[0])
    if len(folder):
        print("清理目录: " + folder)
        print("帅选条件: " + regular)
        y = input("确定清理吗（Y/N）: ")
        print("\n")
        if y == "Y" or y == "y":
            lst = createTreeAsPath(folder, fileRegular = regular, scanSubFolder = True)
            for l in lst:
                os.remove(l)
                print("Removed: " + l)
        os.system("pause")
    root.destroy()

import sys, os
from PIL import Image
from PathTool import createTreeAsPath
from tkinter import Tk, filedialog, messagebox

"""请在此处设置需要转换的图片格式，可以指定多种格式，支持jpg,png,bmp,gif,ico格式，多种格式用'|'分隔"""
fr_type = ".png|.bmp|.gif".lower()  #.jpg|
"""请在此处设置需要转换为的图片格式，支持jpg,png,bmp,gif,ico格式"""
to_type = ".jpg".lower()

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("图片来源", "转换文件夹内所有图片吗？"):
        indir = filedialog.askdirectory(title="请选择需要合并的图片所在的文件夹：",)
        if len(indir):
            infiles = createTreeAsPath(indir, fileRegular = r"^.*("+fr_type+")$", scanSubFolder = False, treeMode = False)
            if not len(infiles):
                messagebox.showinfo("错误", "无符合要求的文件")
                sys.exit(0)
        else:
            messagebox.showinfo("错误", "无符合要求的文件")
            sys.exit(0)
    else:
        if '|' in fr_type:
            temp = ""
            exts = fr_type.split(sep='|')
            for ext in exts:
                if len(temp):
                    temp = temp + ";*" + ext
                else:
                    temp = "*" + ext
            fr_type = temp
        else:
            fr_type = "*" + fr_type
        infiles = filedialog.askopenfilenames(title="请选择需要重置的图片：", filetypes=(("Image File", fr_type), ))
        if not len(infiles):
            messagebox.showinfo("错误", "无符合要求的文件")
            sys.exit(0)
        
    for infile in infiles:
        fname, fexte = os.path.splitext(infile)
        if fexte.lower() != to_type:
            try:
                img = Image.open(infile)
            except:
                messagebox.showinfo("错误","打开失败：" + infile)
            else:
                try:
                    print(fname+to_type)
                    img.save(fname+to_type)
                except:
                    messagebox.showinfo("错误","转换失败：" + infile)
    messagebox.showinfo("提示","任务已完成！" )
    root.destroy()

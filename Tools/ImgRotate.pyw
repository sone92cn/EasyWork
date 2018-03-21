#encoding "utf8"
import os
from PIL import Image
from FelixFunc import resizeRect
from PathTool import createTreeAsPath
from tkinter import Tk, filedialog, messagebox

if __name__ == '__main__':
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("图片来源", "选择文件夹并缩放目录内所有图片吗？"):
        indir = filedialog.askdirectory(title="请选择需要合并的图片所在的文件夹：",)
        if len(indir):
            infiles = createTreeAsPath(indir, fileRegular = r"^.*\.(jpg|jpeg|png)$", scanSubFolder = False, treeMode = False)
        else:
            infiles = []
    else:
        infiles = filedialog.askopenfilenames(title="请选择需要重置的图片：", filetypes=(("Image Files", "*.jpg;*.jpeg;*.png"), ))
    
    if len(infiles):
        overwrite =  messagebox.askyesno("重置类型", "覆盖原图保存吗？")
        for file in infiles:
            img = Image.open(file)  
            img = img.transpose(Image.ROTATE_180) #旋转180度角。
            if overwrite:
                img.save(file)
            else:
                img.save(os.path.join(os.path.split(file)[0], 'MD_' + os.path.split(file)[1]))
        messagebox.showinfo("提示", "已完成！")
    else:
        messagebox.showerror("Error", "No Image(s) selected!")
    root.destroy()

import os
from FelixFunc import resizeRect
from tkinter import Tk, filedialog, messagebox
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import landscape, A4

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    infiles = filedialog.askopenfilenames(title="请选择需要添加到PDF的图片：", filetypes=(("Image File", "*.jpg;*.jpeg;*.png;*.bmp"),))
    if len(infiles):
        errs = ""
        land = messagebox.askyesno(title="选择", message="页面是否为纵向？")
        for infile in infiles:
            outfile = os.path.splitext(infile)[0] + ".pdf"
            try:
                if land:
                    cas = Canvas(outfile, pagesize=A4)
                else:
                    cas = Canvas(outfile, pagesize=landscape(A4))
            except:
                errs = errs + "\n" + outfile
            else:
                s_w, s_h = s_size = cas._pagesize
                img = ImageReader(infile)
                i_w, i_h = resizeRect(img.getSize(), s_size)
                cas.drawImage(img, (s_w-i_w)/2, (s_h-i_h)/2, i_w, i_h)
                cas.save()
        if len(errs):
            messagebox.showinfo("","转换失败：" + outfile)
        else:
            messagebox.showinfo("","任务已完成！")
    root.destroy()
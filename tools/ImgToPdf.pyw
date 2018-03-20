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
        if messagebox.askyesno(title="选择", message="是否逆序添加？"):
            infiles = reversed(infiles)
        else:
            infiles = sorted(infiles)
        outfile = filedialog.asksaveasfilename(title="保存至：", filetypes=(("PDF File", "*.pdf"),))
        if len(outfile):
            if os.path.splitext(outfile)[1] != '.pdf':
                outfile += '.pdf'
            try:
                if messagebox.askyesno(title="选择", message="页面是否为纵向？"):
                    cas = Canvas(outfile, pagesize=A4)
                else:
                    cas = Canvas(outfile, pagesize=landscape(A4))
            except:
                messagebox.showinfo('ERROR', '无法打开：' + outfile)
            else:
                new = False
                s_w, s_h = s_size = cas._pagesize
                for infile in infiles:
                    if new:
                        cas.showPage()
                    else:
                        new = True
                    img = ImageReader(infile)
                    i_w, i_h = resizeRect(img.getSize(), s_size)
                    cas.drawImage(img, (s_w-i_w)/2, (s_h-i_h)/2, i_w, i_h)
                cas.save()
                messagebox.showinfo("","成功保存到：" + outfile)
    root.destroy()

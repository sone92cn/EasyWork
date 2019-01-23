import os
from PathTool import createTreeAsPath
from tkinter import Tk, filedialog, messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("数据来源", "选择文件夹并旋转目录内所有PDF文件吗？"):
        indir = filedialog.askdirectory(title="请选择需要旋转的PDF所在的文件夹：",)
        if len(indir):
            infiles = createTreeAsPath(indir, fileRegular = '^.*\.pdf$', scanSubFolder = False, treeMode = False)
        else:
            infiles = []
    else:
        infiles = filedialog.askopenfilenames(title="请选择需要旋转的文件：", filetypes=(("PDF Files", "*.pdf"), ))

    i, n = 0, 0
    for infile in infiles:
        try:
            outpdf = PdfFileWriter()
            inpdf = PdfFileReader(open(infile, "rb"))
            fpath, fname = os.path.split(infile)
            npath = os.path.join(fpath, "new")
            if not os.path.isdir(npath):
                os.mkdir(npath)
            for page in inpdf.pages:
                outpdf.addPage(page.rotateClockwise(180))
            outStream = open(os.path.join(npath, fname), "wb")
            outpdf.write(outStream)
            outStream.close()
        except:
            i += 1
        else:
            n += 1
    messagebox.showinfo("","成功%d个文件，失败%d个文件。" % (n, i))
            
    root.destroy()

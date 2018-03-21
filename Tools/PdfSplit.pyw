import os
from FelixFunc import getStrFromInt
from tkinter import Tk, filedialog, messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    infile = filedialog.askopenfilename(title="请选择需要拆分的PDF文件：", filetypes=(("PDF Files", "*.pdf"), ))
    if len(infile):
        path, file = os.path.split(infile)
        name, exte = os.path.splitext(file)
        outpath = filedialog.askdirectory(title="请指定要另存到的文件夹：", initialdir=path)
        if len(outpath):
            inpdf = PdfFileReader(open(infile, "rb"))
            outpdfs = []
            for page in inpdf.pages:
                outpdf = PdfFileWriter()
                outpdf.addPage(page)
                outpdfs.append(outpdf)
            i = 1
            print(outpath)
            for outpdf in outpdfs:
                outfile = os.path.join(outpath, name+getStrFromInt(i, 3)+exte)
                while os.path.isfile(outfile):
                    i += 1
                    outfile = os.path.join(outpath, name+getStrFromInt(i, 3)+exte)
                print(outfile)
                outStream = open(outfile, "wb")
                outpdf.write(outStream)
                outStream.close()
            messagebox.showinfo("","成功保存到：" + outpath)
    root.destroy()
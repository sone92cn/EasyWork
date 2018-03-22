from PathTool import createTreeAsPath
from tkinter import Tk, filedialog, messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    if messagebox.askyesno("数据来源", "选择文件夹并合并目录内所有PDF文件吗？"):
        indir = filedialog.askdirectory(title="请选择需要合并的PDF所在的文件夹：",)
        if len(indir):
            infiles = createTreeAsPath(indir, fileRegular = '^.*\.pdf$', scanSubFolder = False, treeMode = False)
        else:
            infiles = []
    else:
        infiles = filedialog.askopenfilenames(title="请选择需要合并的文件：", filetypes=(("PDF Files", "*.pdf"), ))
        
    if len(infiles):
        outfile = filedialog.asksaveasfilename(title="请指定要另存为的文件：", filetypes=(("PDF Files", "*.pdf"), ))
        if len(outfile):
            if outfile[-4:] != ".pdf":
                outfile += ".pdf"
            outpdf = PdfFileWriter()
            infiles = sorted(infiles)  #集合只支持次方式 infiles.sort()此处不能用
            for infile in infiles:
                if infile != outfile:
                    inpdf = PdfFileReader(open(infile, "rb"))
                    for page in inpdf.pages:
                        outpdf.addPage(page)
            outStream = open(outfile, "wb")
            outpdf.write(outStream)
            outStream.close()
            messagebox.showinfo("","成功保存至" + outfile)
    root.destroy()

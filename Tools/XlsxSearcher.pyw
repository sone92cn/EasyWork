#encoding "UTF-8"
import os, re, time, threading 
from PathTool import createTreeAsPath
from FelixFunc import wipeNone
from ExcelTool import ExcelReader, ExcelWriter
from ThreadTool import MissionEx, ManagerEx
from tkinter import Tk, ttk, Listbox, filedialog, messagebox, BooleanVar, IntVar, StringVar

class staticMethod():
    @staticmethod
    def main():
        thinterWindow("Excel数据搜索").showWindow()
    @staticmethod
    def loadFiles(args):
        if len(args) == 7:
            fname, group, vdict, failed, mission, lock, shets = args
            try:
                vdict["var_info"].set('Loading: ' + fname)
                xls = ExcelReader(fname, group)
            except:
                lock.acquire()
                failed.append(fname)
                lock.release()
                vdict["var_info"].set('Fail to load: ' + fname)
            else:
                vdict["var_info"].set('Succeed to load: ' + fname)
                for sheet in xls.valid_sheets:
                    shet = ExcelReader.getGroupShet(group, sheet)
                    mission.addMission(shet)
                    shets.append(shet)
    @staticmethod
    def searchShets(args):
        if len(args) == 6:
            """(int(text), text)=infos"""
            shet, stype, smode, sinfos, nmode, vdict = args
            saves = []
            nrows = shet['xls'].getTotalRows(shet['sheet'])
            ncols = shet['xls'].getTotalCols(shet['sheet'])
            if not('shet_info' in shet):
                shet['shet_info'] = shet['xls'].getSheetName(shet['sheet']) + '<' + shet['xls'].getBookName() + '>'
            vdict["var_info"].set('Searching: ' + shet['shet_info'])
            for row in range(nrows):
                for col in range(ncols):
                    info = shet['xls'].getCellValue(shet['sheet'], row, col, check_merge=True, return_top=True, collect_all=True)
                    if info.type == stype:
                        if smode == 'eq':
                            for sinfo in sinfos:
                                if sinfo[0] == info.value:
                                    saves.append((shet, row, sinfo[1]))
                                    break
                        elif smode == 'fd':
                            for sinfo in sinfos:
                                if sinfo[0] in info.value:
                                    saves.append((shet, row, sinfo[1]))
                                    break
                        elif smode == 'mh':
                            for sinfo in sinfos:
                                if sinfo[0].match(info.value) != None:
                                    saves.append((shet, row, sinfo[1]))
                                    break
            if len(saves):
                if not('title_row' in shet):
                    try:
                        shet['title_row'] = shet['xls'].guessTitleRow(shet['sheet'], allow_jump=True)
                    except:
                        shet['title_row'] = None          
                s_names = list(set([save[2] for save in saves]))
                for s_name in s_names:
                    s_objs = []
                    s_rows = []
                    writers = ExcelWriter.getWriter(s_name, nmode)
                    for save in saves:
                        if save[2] == s_name:
                            if save[1] in s_rows:
                                continue
                            else:
                                s_objs.append(save)
                                s_rows.append(save[1])
                    for writer in writers:
                        writer['lock'].acquire()
                        if shet['title_row'] == None:
                            writer['xls'].appendRowValues(writer['sheet'], ['Fail to Guess Title', 'from ' + shet['shet_info']])
                        else:
                            vals = shet['xls'].getRowValues(shet['sheet'], shet['title_row'])
                            vals.append('from ' + shet['shet_info'])
                            writer['xls'].appendRowValues(writer['sheet'], vals)
                        for s_obj in s_objs:
                            writer['xls'].appendRowValues(writer['sheet'], s_obj[0]['xls'].getRowValues(s_obj[0]['sheet'], s_obj[1]))
                        writer['lock'].release()
                        
    @staticmethod
    def chooseKeyFile(proc, vdict):
        open_file = filedialog.askopenfilename(title="选择文件", initialdir=proc.workdir, filetypes=(("Excel Files", "*.xls;*.xlsx;*.xlsm"), ))
        if len(open_file):
            vdict["var_info"].set('Loading: ' + open_file)
            proc.run = True
            xls = ExcelReader(open_file, group='info')
            for i, sheet in enumerate(xls.valid_sheets):
                try:
                    r = xls.guessTitleRow(sheet, allow_jump=True)
                except:
                    continue
                else:
                    staticMethod.resetCombobox(vdict['check_sheet'], xls.valid_sheet_names, i)
                    vals = [i+1 for i in range(9)]
                    staticMethod.resetCombobox(vdict['check_title'], vals, r)
                    staticMethod.resetCombobox(vdict['check_column'], list(set(wipeNone(xls.getRowValues(sheet, r)))))
                    break
            else:
                staticMethod.resetCombobox(vdict['check_sheet'], xls.valid_sheet_names, 0)
                vals = [i+1 for i in range(9)]
                staticMethod.resetCombobox(vdict['check_title'], vals, 0)
                staticMethod.resetCombobox(vdict['check_column'], list(set(wipeNone(xls.getRowValues(sheet, 0)))))
            proc.run = False
            proc.xls = xls
            proc.detailChanged(vdict)
            vdict['frame_details'].grid(row=0, column=2, rowspan=2)
            vdict['var_from'].set(1)
            vdict['var_info'].set('Loaded: ' + open_file)
            time.sleep(5)
            vdict['var_info'].set('')
        else:
            vdict['frame_details'].grid_forget()
            vdict['var_from'].set(0)
            vdict['check_sheet'].set('')
            vdict['check_title'].set('')
            vdict['check_column'].set('')
            vdict['check_sheet']['values'], vdict['check_title']['values'], vdict['check_column']['values'] = [], [], []
    @staticmethod
    def resetListBox(box, data, maxlen=0):
        box.delete(0, "end")
        for d in data:
            if maxlen:
                if len(d) > maxlen:
                    box.insert("end", '..' + d[2-maxlen:])
                else:
                    box.insert("end", d)
            else:
                box.insert("end", d)
    @staticmethod
    def resetCombobox(cbx, shnames, current=0):
        cbx['values'] = shnames
        cbx.current(current)
                
class tkinterProcdure():
    def __init__(self, hwnd):
        self.run = 0
        self.hwnd = hwnd
        self.mdict = {}
        self.workdir = ''
        self.xls = None
        self.shets = []
        self.smode = 0
        self.mission = None
        self.ire = re.compile('^[0-9]+$')
        self.fre = re.compile('^[0-9]+\.[0-9]+$')
    def __getSCodes(self, mode):
        if mode == 1:
            return ('no', 'no')
        elif mode == 2: #数值
            return ('n', 'eq')
        elif mode == 3: #精确文本
            return ('s', 'eq')
        elif mode == 4: #模糊文本
            return ('s', 'fd')
        elif mode ==5:  #正则表达式
            return ('s', 'mh')
        else:
            return ('no', 'no')
    def __getCheckDetail(self, text, mode):
        """return None or (s_type, s_mode, s_value, s_text)"""
        if mode == 1:
            return None
        elif mode == 2: #数值
            if self.ire.match(text) != None:
                return (int(text), text)
            elif self.fre.match(text) != None:
                return (float(text), text)
            else:
                return None
        elif mode == 3: #精确文本
            return (text, text)
        elif mode == 4: #模糊文本
            return (text, text)
        elif mode ==5:  #正则表达式
            try:
                regu = re.compile(text, re.IGNORECASE)
            except:
                return None
            else:
                return (regu, text)
        else:
            return None
    def sheetChanged(self, vdict):
        sheet = self.xls.getSheetObject(vdict['check_sheet'].get())
        try:
            r = self.xls.guessTitleRow(sheet)
        except:
            vals = [i+1 for i in range(9)]
            staticMethod.resetCombobox(vdict['check_title'], vals, 0)
            staticMethod.resetCombobox(vdict['check_column'], list(set(wipeNone(self.xls.getRowValues(sheet, 0)))))
        else:
            vals = [i+1 for i in range(9)]
            staticMethod.resetCombobox(vdict['check_title'], vals, r)
            staticMethod.resetCombobox(vdict['check_column'], list(set(wipeNone(self.xls.getRowValues(sheet, r)))))
        self.detailChanged(vdict)
    def titleChanged(self, vdict):
        sheet = self.xls.getSheetObject(vdict['check_sheet'].get())
        trow = vdict['check_title'].current()
        staticMethod.resetCombobox(vdict['check_column'], list(set(wipeNone(self.xls.getRowValues(sheet, trow)))))
        self.detailChanged(vdict)
    def detailChanged(self, vdict):
        sheet = self.xls.getSheetObject(vdict['check_sheet'].get())
        trow = vdict['check_title'].current()
        col = self.xls.getRowValues(sheet, trow).index(vdict['check_column'].get())
        staticMethod.resetListBox(vdict['listbox_details'], set(self.xls.getColValues(sheet, col, trow+1)))
    def chooseDataFiles(self, vdict):
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        if messagebox.askyesno("选择", "直接指定文件吗？"):
            open_files = filedialog.askopenfilenames(title="选择文件", initialdir=self.workdir, filetypes=(("Excel Files", "*.xls;*.xlsx;*.xlsm"), ))
        else:
            open_dir = filedialog.askdirectory(title="选择文件夹", initialdir = self.workdir)
            if len(open_dir):
                open_files = createTreeAsPath(open_dir, fileRegular=r"^(?!~).*\.(xls|xlsx|xlsm)$", scanSubFolder=vdict['var_scansub'].get(), treeMode=False)
        if len(open_files):
            self.shets = []
            self.mdict['files'] = open_files
            staticMethod.resetListBox(vdict['listbox_files'], open_files, 20)
            
            s_list = []
            self.mission = MissionEx(s_list, 0)
            
            v_list = []  
            v_list.extend(open_files)
            """fname, group, vdict, failed, mission, lock = args"""
            mission = MissionEx(v_list, 0, staticMethod.loadFiles, args=(None, 'reader', vdict, [], self.mission, threading.Lock(), self.shets))
            ManagerEx(mission, 4, block_primary=True, new_thread=True, end_call=self.fileLoaded).start()
    def chooseKeyFile(self, vdict, value):
        if value:
            threading.Thread(target=staticMethod.chooseKeyFile, args=(self, vdict)).start()
        else:
            vdict['frame_details'].grid_forget()
            vdict['var_from'].set(0)
            vdict['check_sheet'].set('')
            vdict['check_title'].set('')
            vdict['check_column'].set('')
            vdict['check_sheet']['values'], vdict['check_title']['values'], vdict['check_column']['values'] = [], [], []
    def confirmCommand(self, vdict):
        if self.run:
            messagebox.showinfo('ERROR', '正在搜索，请耐心等待！')
            return
        if self.mission == None:
            messagebox.showinfo('ERROR', '请选择将要搜索的文件！')
            return
        mode = vdict['var_mode'].get()
        if vdict['var_from'].get():
            texts = vdict['listbox_details'].get(0, 'end')
            s_infos = []
            for text in texts:
                s_info = self.__getCheckDetail(text, mode)
                if s_info == None:
                    messagebox.showinfo('ERROR', '搜索的方式与内容不完全匹配！')
                    return False
                else:
                    s_infos.append(s_info)
            self.smode = vdict['var_smode'].get()
        else:
            s_infos = []
            text = vdict['var_text'].get()
            if text == '':
                messagebox.showinfo('ERROR', '请输入要搜索的内容！')
                return False
            s_info = self.__getCheckDetail(text, mode)
            if s_info == None:
                messagebox.showinfo('ERROR', '搜索的方式与内容不匹配！')
                return False
            s_infos.append(s_info)
            self.smode = 0
        
        ExcelWriter.resetBook()
        ExcelWriter.resetWriters()
        stype, smode = self.__getSCodes(mode)
        if not self.mission.valid:
            vlist = []
            vlist.extend(self.shets)
            self.mission.resetVarList(vlist)
        """shet, stype, smode, sinfos, nmode = args"""
        self.mission.createMission(staticMethod.searchShets, args=(None, stype, smode, s_infos, self.smode, vdict))
        manager = ManagerEx(self.mission, 4, block_primary=True, new_thread=True, end_call=self.searchOver)
        self.run = True
        manager.start() 
    def fileLoaded(self, success, msg, args):
        self.hwnd.vdict['var_info'].set('All files loaded!')
        self.mission.endMission()
        if not self.run:
            time.sleep(5)
            self.hwnd.vdict['var_info'].set('')
    def searchOver(self, success, msg, args): 
        self.hwnd.vdict['var_info'].set('Search finished!')
        if self.smode:
            folder = filedialog.askdirectory(title="结果储存到：")
            if len(folder):
                shets = [shet for writer in ExcelWriter.writers if writer != 'lock' for shet in ExcelWriter.writers[writer]] 
                xfs = ExcelWriter.getExcelObjects(shets)
                if len(xfs):
                    for xf in xfs:
                        fname = xf.getBookName()
                        if len(fname) == 0:
                            fname = xf.file
                        try:
                            xf.saveas(os.path.join(folder, fname))
                        except:
                            messagebox.showinfo("Error", "储存搜索结果失败：" + xf.getBookName())
        else:
            try:
                ExcelWriter.saveBook()
            except:
                messagebox.showinfo("Error", "储存搜索结果失败!")
        self.run = False
        self.hwnd.vdict["var_info"].set('Mission Over')
        time.sleep(3)
        self.hwnd.vdict['var_info'].set('')
    
class thinterWindow():
    def __init__(self, title):
        self.root = Tk()
        self.proc = tkinterProcdure(self)
        self.vdict = {'name':'cmd'}
        self.__createLeftWidgets(self.vdict).grid(row=0, column=0, sticky=('nsew'), rowspan=2)
        self.__createRightWidgets(self.vdict).grid(row=0, column=1, sticky=('nsew'))
        self.__createConfirmWidgets(self.vdict).grid(row=1, column=1, sticky=('ew'))
        self.__createFootWidgets(self.vdict).grid(row=2, column=0, sticky=('w'), columnspan=3)
        self.vdict['frame_details'] = self.__createDetailWidgets(self.vdict)
        self.__configureWindow(title)
    def __createLeftWidgets(self, vdict):
        """create frame"""
        frame = ttk.Frame(self.root, padding="8")
        """create variables"""
        scansub = BooleanVar(value=True)
        """create components"""
        lab00 = ttk.Label(frame, text='请选择要搜索的文件：')
        lst10 = Listbox(frame, exportselection=False)
        sib12 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        cbn30 = ttk.Checkbutton(frame, variable=scansub, onvalue=True, offvalue=False, text="扫描子目录")
        """grid components"""
        lab00.grid(row=0, column=0, sticky=('w'))
        lst10.grid(row=1, column=0, columnspan=2, sticky=('w', 'e'))
        sib12.grid(row=1, column=2, sticky=('s', 'n', 'w'))
        sib20.grid(row=2, column=0, columnspan=2, sticky=('w', 'n', 'e'))
        cbn30.grid(row=3, column=0, sticky=('w'))
        """save components"""
        vdict['label_choose'] = lab00
        vdict['listbox_files'] = lst10
        vdict['var_scansub'] = scansub
        return frame
    def __createRightWidgets(self, vdict):
        """create frame"""
        frame = ttk.Frame(self.root, padding="8")
        """create variables"""
        tmode, tfrom, stext, smode = IntVar(value=3), IntVar(value=0), StringVar(), IntVar(value=0)
        """create and grid components"""
        frame1 = ttk.Frame(frame, padding="0")
        lab00 = ttk.Label(frame1, text='搜索方式：')
        lab00.grid(row=0, column=0, sticky=('w'))
        frame1.grid(row=0, column=0, sticky=('nsew'))
        frame2 = ttk.Frame(frame, padding="0")
        rbn10 = ttk.Radiobutton(frame2, variable=tmode, text='日期', value=1, state='disabled')
        rbn11 = ttk.Radiobutton(frame2, variable=tmode, text='数值', value=2)
        rbn20 = ttk.Radiobutton(frame2, variable=tmode, text='精确文本', value=3)
        rbn21 = ttk.Radiobutton(frame2, variable=tmode, text='模糊文本', value=4)
        rbn22 = ttk.Radiobutton(frame2, variable=tmode, text='正则表达式', value=5)
        rbn10.grid(row=0, column=0, sticky=('w'))
        rbn11.grid(row=0, column=1, sticky=('w'))
        rbn20.grid(row=1, column=0, sticky=('w'))
        rbn21.grid(row=1, column=1, sticky=('w'))
        rbn22.grid(row=1, column=2, sticky=('w'))
        frame2.grid(row=1, column=0, sticky=('nsew'))
        frame3 = ttk.Frame(frame, padding="0")
        lab20 = ttk.Label(frame3, text='内容来源：')
        lab20.grid(row=0, column=0, sticky=('w'))
        frame3.grid(row=2, column=0, sticky=('nsew'))
        frame4 = ttk.Frame(frame, padding="0")
        rbn30 = ttk.Radiobutton(frame4, variable=tfrom, text='手工输入', value=0)
        ent31 = ttk.Entry(frame4, textvariable=stext)
        rbn40 = ttk.Radiobutton(frame4, variable=tfrom, text='来自文件', value=1)
        lab50 = ttk.Label(frame4, text='工作表：')
        cbx51 = ttk.Combobox(frame4, width=12, state="readonly")
        lab60 = ttk.Label(frame4, text='标题行：')
        cbx61 = ttk.Combobox(frame4, width=12, state="readonly")
        lab70 = ttk.Label(frame4, text='数据列：')
        cbx71 = ttk.Combobox(frame4, width=12, state="readonly")
        rbn30.grid(row=0, column=0, sticky=('w'))
        ent31.grid(row=0, column=1, sticky=('w'))
        rbn40.grid(row=1, column=0, sticky=('w'), columnspan=2)
        lab50.grid(row=2, column=0, sticky=('e'))
        cbx51.grid(row=2, column=1, sticky=('w'))
        lab60.grid(row=3, column=0, sticky=('e'))
        cbx61.grid(row=3, column=1, sticky=('w'))
        lab70.grid(row=4, column=0, sticky=('e'))
        cbx71.grid(row=4, column=1, sticky=('w'))
        frame4.grid(row=3, column=0, sticky=('nsew'))
        frame5 = ttk.Frame(frame, padding="0")
        lab80 = ttk.Label(frame5, text='储存到：')
        rbn90 = ttk.Radiobutton(frame5, variable=smode, text='单个文件', value=0)
        rbn91 = ttk.Radiobutton(frame5, variable=smode, text='多个文件', value=1)
        lab80.grid(row=0, column=0, sticky=('w'))
        rbn90.grid(row=0, column=1, sticky=('w'))
        rbn91.grid(row=0, column=2, sticky=('w'))
        frame5.grid(row=4, column=0, sticky=('nsew'))
        """save components"""
        vdict['var_mode'] = tmode
        vdict['var_from'] = tfrom
        vdict['var_text'] = stext
        vdict['var_smode'] = smode
        vdict['radio_text'] = rbn30
        vdict['radio_file'] = rbn40
        vdict['check_sheet'] = cbx51
        vdict['check_title'] = cbx61
        vdict['check_column'] = cbx71
        return frame
    def __createConfirmWidgets(self, vdict):
        frame = ttk.Frame(self.root, padding="76 8 72 0")   #padding   W N E S
        btn00 = ttk.Button(frame, text='搜  索')
        btn00.grid(row=0, column=0, sticky=('nsew'))
        vdict['button_confirm'] = btn00
        return frame
    def __createFootWidgets(self, vdict):
        frame = ttk.Frame(self.root, padding="0")
        labinfo = StringVar()
        lab00 = ttk.Label(frame, textvariable=labinfo, text='')
        lab00.grid(row=0, column=0, sticky=('w'))
        vdict['var_info'] = labinfo
        return frame
    def __createDetailWidgets(self, vdict):
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text='将搜索以下内容：')
        lst10 = Listbox(frame, exportselection=False)
        sib12 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        lab00.grid(row=0, column=0, sticky=('w'))
        lst10.grid(row=1, column=0, columnspan=2, sticky=('w', 'e'))
        sib12.grid(row=1, column=2, sticky=('s', 'n', 'w'))
        sib20.grid(row=2, column=0, columnspan=2, sticky=('w', 'n', 'e'))
        vdict['listbox_details'] = lst10
        return frame
    def __configureWindow(self, title):
        self.root.title(title)
        self.vdict['label_choose'].bind('<1>', lambda e: self.proc.chooseDataFiles(self.vdict))
        self.vdict['radio_text'].bind('<1>', lambda e: self.proc.chooseKeyFile(self.vdict, 0))
        self.vdict['radio_file'].bind('<1>', lambda e: self.proc.chooseKeyFile(self.vdict, 1))
        self.vdict['check_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.vdict))
        self.vdict['check_title'].bind('<<ComboboxSelected>>', lambda e: self.proc.titleChanged(self.vdict))
        self.vdict['check_column'].bind('<<ComboboxSelected>>', lambda e: self.proc.detailChanged(self.vdict))
        self.vdict['button_confirm'].bind('<1>', lambda e: self.proc.confirmCommand(self.vdict))
    def destroyWindow(self):
        self.root.destroy()
    def showWindow(self):
        self.root.mainloop()
        
if __name__ == "__main__":
    staticMethod.main()
#encoding "UTF-8"
import os, threading, time
from PathTool import createTreeAsPath
from ThreadTool import MissionEx, ManagerEx
from ExcelTool import ExcelReader, ExcelWriter
from FelixFunc import combineLists, wipeNone
from tkinter import Tk, Toplevel, ttk, Listbox, Text, filedialog, messagebox, BooleanVar, StringVar
    
class staticMethod():
    @staticmethod
    def main():
        thinterWindow().showMainWindow()
    @staticmethod
    def loadFile(args):
        if len(args) == 5:
            fname, group, read_only, info_fict, failed_to = args
            try:
                if read_only:
                    xls = ExcelReader(fname, group)
                else:
                    xls = ExcelWriter(fname, group)
            except:
                failed_to.append(fname)
                info_fict["var_information"].set('Fail to load: ' + fname)
                return None
            else:
                info_fict["var_information"].set('Succeed to load: ' + fname)
                return xls
        else:
            raise Exception('Args error!')
    @staticmethod
    def checkTitle(args):
        if len(args) == 7:
            shet, check_list, key_len, guess_row, max_row, cls_obj, info_dict = args
            info_dict["var_information"].set('Checking: ' + shet['info'])
            try:
                check_pos = []
                row = shet['xls'].getTitleRow(shet['sheet'], check_list, check_pos, guess_row, max_row)
            except:
                shet['lock'].acquire()
                shet['valid'] = False
                shet['lock'].release()
                cls_obj.groups['lock'].acquire()
                cls_obj.groups['invalid'].append(shet)
                cls_obj.groups['lock'].release()
            else:
                shet['lock'].acquire()
                shet['valid'] = True
                shet['title_row'] = row
                shet['key_cols'] = check_pos[:key_len]
                shet['compare_cols'] = check_pos[key_len:]
                shet['total_col'] = shet['xls'].getTotalCols(shet['sheet'])
                shet['lock'].release()
            info_dict["var_information"].set('Checked: ' + shet['info'])
        else:
            raise Exception('Args error!')
    @staticmethod
    def checkKeys(args):
        """getRowsValuesByPos(sheet, pos, start_row=0, check_merge=True, return_top=True, collect_all=False, skip_none=False, skip_allnone=True, skip_sumrow=True)"""
        if len(args) == 3:
            shet, not_null, info_dict = args
            info_dict["var_information"].set('Preparing: ' + shet['info'])
            key_list = shet['xls'].getRowsValuesByPos(shet['sheet'], shet['key_cols'], start_row=shet['title_row']+1, skip_none=not_null)
            shet['lock'].acquire()
            shet['keys'] = key_list
            shet['lock'].release()
            info_dict["var_information"].set('Prepared: ' + shet['info'])
        else:
            raise Exception('Args error!')
    @staticmethod
    def checkFrom(args):    #search all data shets and auto append
        if len(args) == 3:
            mshet, dshets, info_dict = args
            mshet['lock'].acquire()
            mkeys = mshet['keys']
            info_dict["var_information"].set('Checking: ' + mshet['info'])
            i = 0
            print('mkeys:', mkeys)
            for mrow in mkeys:
                for dshet in dshets:
                    dkeys = dshet['keys']
                    if i < 2:
                        print('dkeys:', dkeys)
                        i += 1
                    for drow in dkeys:
                        if mkeys[mrow] == dkeys[drow]:
                            if mshet['xls'].getRowValuesByPos(mshet['sheet'], mrow, mshet['compare_cols'], check_merge=False) == dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['compare_cols'], check_merge=False):
                                mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], '验证相符：' + dshet['info'])
                            else:
                                print(mshet['xls'].getRowValuesByPos(mshet['sheet'], mrow, mshet['compare_cols'], check_merge=False), dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['compare_cols'], check_merge=False))
                                mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], '验证不符：' + dshet['info'])
                            break
                    else:
                        continue
                    break
                else:
                    writers = ExcelWriter.getWriter('未搜索到的数据', new_mode=0)
                    for writer in writers:
                        writer['lock'].acquire()
                        nrow = writer['xls'].getTotalRows(writer['sheet'])
                        writer['xls'].updateRowValues(writer['sheet'], nrow, mshet['xls'].getRowValues(mshet['sheet'], mrow))
                        writer['xls'].updateCellValue(writer['sheet'], nrow, mshet['total_col'], 'FROM' + mshet['info'])
                        writer['lock'].release()
            mshet['lock'].release()
        else:
            raise Exception('Args Error!')
    @staticmethod
    def checkTo(args):    #search all data shets and auto append
        if len(args) == 3:
            dshet, mshets, info_dict = args
            dkeys = dshet['keys']
            info_dict["var_information"].set('Checking: ' + dshet['info'])
            for drow in dkeys:
                for mshet in mshets:
                    mshet['lock'].acquire()
                    mkeys = mshet['keys']
                    for mrow in mkeys:
                        if mkeys[mrow] == dkeys[drow]:
                            if mshet['xls'].getRowValuesByPos(mshet['sheet'], mrow, mshet['compare_cols'], check_merge=False) == dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['compare_cols'], check_merge=False):
                                mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], '验证相符：' + dshet['info'])
                            else:
                                mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], '验证不符：' + dshet['info'])
                            break
                    else:
                        mshet['lock'].release()
                        continue
                    mshet['lock'].release()
                    break
                else:
                    writers = ExcelWriter.getWriter('仅存在于检查依据的记录', new_mode=0)
                    for writer in writers:
                        writer['lock'].acquire()
                        nrow = writer['xls'].getTotalRows(writer['sheet'])
                        writer['xls'].updateRowValues(writer['sheet'], nrow, dshet['xls'].getRowValues(dshet['sheet'], drow))
                        writer['xls'].updateCellValue(writer['sheet'], nrow, dshet['total_col'], 'FROM' + dshet['info'])
                        writer['lock'].release()
        else:
            raise Exception('Args error!')
    @staticmethod
    def clearListBoxes(boxes):
        for box in boxes:
            box.delete(0, 'end')
    @staticmethod
    def clearComboboxes(boxes):
        for box in boxes:
            box.set('')
            box['values'] = []
    @staticmethod
    def resetCombobox(cbx, shnames, current):
        cbx['values'] = shnames
        cbx.current(current)
    @staticmethod
    def resetRowCombobox(cbx, mxrow, tlrow):
        cbx['values'] = [row+1 for row in range(mxrow+1)]
        cbx.current(tlrow)
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
    def resetListBoxes(xls, sheet, tlrow, boxes):
        items = wipeNone(xls.getRowValues(sheet, tlrow, check_merge=False))
        if len(items) == len(set(items)):
            for box in boxes:
                box.delete(0, 'end')
            for item in items:
                for box in boxes:
                    box.insert('end', item)
        else:
            staticMethod.clearListBoxes(boxes)
            messagebox.showinfo("Error", "标题行不可重复!")
    @staticmethod
    def autoSelectListboxes(vdict):
        bdict = vdict['brother']
        for lname in ['listbox_keys', 'listbox_compare']:
            selected_items = staticMethod.getSelectedListboxItems(bdict[lname], bdict[lname].curselection())
            for item in selected_items:
                staticMethod.selectListBoxItemByValue(vdict[lname], item, True)
    @staticmethod
    def selectListBoxItemByValue(box, item, action):
        for i in range(box.size()):
            if box.get(i) == item:
                if action:
                    if not box.selection_includes(i):
                        box.selection_set(i)
                        return i
                else:
                    if box.selection_includes(i):
                        box.selection_clear(i)
                        return i
                break
        return -1
    @staticmethod
    def saveSelectedItems(mdict, vdict):
        mbdict, vbdict = mdict['brother'], vdict['brother']
        mdict['listbox_keys'] = vdict['listbox_keys'].curselection()
        mdict['listbox_compare'] = vdict['listbox_compare'].curselection()
        mbdict['listbox_keys'] = vbdict['listbox_keys'].curselection()
        mbdict['listbox_compare'] = vbdict['listbox_compare'].curselection()
    @staticmethod
    def getSelectedListboxItems(box, selected):
        items = []
        for i in selected:
            items.append(box.get(i))
        return items
    @staticmethod
    def getValidateSheetDict(vdict, mdict, idict):
        key = staticMethod.getSelectedListboxItems(vdict['listbox_keys'], mdict['listbox_keys'])
        compare = staticMethod.getSelectedListboxItems(vdict['listbox_compare'], mdict['listbox_compare'])
        total = combineLists(key, compare)
            
        if mdict['name'] == 'main':
            if vdict['var_onlyone'].get():
                shets = [ExcelWriter.getGroupShet('writer', mdict['xls'].getSheetObject(vdict['combobox_sheet'].get()))]
            else:
                shets = ExcelWriter.groups['writer']
            ExcelWriter.resetGroup('invalid')
            """shet, check_list, key_len, guess_row, max_row, cls_obj, info_dict = args"""
            missions = MissionEx(shets, 0, staticMethod.checkTitle, (None, total, len(key), vdict['combobox_row'].current(), 9, ExcelWriter, idict))
        else:
            if vdict['var_onlyone'].get():
                shets = [ExcelReader.getGroupShet('reader', mdict['xls'].getSheetObject(vdict['combobox_sheet'].get()))]
            else:
                shets = ExcelReader.groups['reader']
            ExcelReader.resetGroup('invalid')
            """shet, check_list, key_len, guess_row, max_row, cls_obj, info_dict = args"""
            missions = MissionEx(shets, 0, staticMethod.checkTitle, (None, total, len(key), vdict['combobox_row'].current(), 9, ExcelReader, idict))
        ManagerEx(missions, block_primary=True, new_thread=False).start()
        return [shet for shet in shets if shet['valid']]
    @staticmethod
    def confirmCommand(obj, vdict):
        obj.run = True
        mdict_main = obj.main_dict
        mdict_data = obj.data_dict
        vdict_main = obj.app.main_dict
        vdict_data = obj.app.data_dict
        staticMethod.saveSelectedItems(mdict_main, vdict_main)
        valid_main = staticMethod.getValidateSheetDict(vdict_main, mdict_main, vdict)
        valid_data = staticMethod.getValidateSheetDict(vdict_data, mdict_data, vdict)
        if len(ExcelReader.groups['invalid']) or len(ExcelWriter.groups['invalid']):
            stop = BooleanVar(value=True)
            conf = confirmWindow(obj.app.root, stop)
            obj.app.root.wait_window(conf.root)
            if stop.get():
                obj.run = False
                time.sleep(3)
                vdict["var_information"].set('')
                return
        missions = MissionEx(combineLists(valid_main, valid_data), 0, staticMethod.checkKeys, (None, vdict_main['var_notnull'].get(), vdict))
        ManagerEx(missions, block_primary=True, new_thread=False).start()
        
        ExcelWriter.resetBook()
        ExcelWriter.resetWriters()
        if vdict['var_allmatch'].get():
            missions = MissionEx(valid_data, 0, staticMethod.checkTo, args=(None, valid_main, vdict))
            ManagerEx(missions, block_primary=True, new_thread=False).start()
            writers = ExcelWriter.getWriter('仅存在于检查数据的记录', new_mode=0)
            vdict["var_information"].set('Collecting results ...')
            for writer in writers:
                writer['lock'].acquire()
                for mshet in valid_main:
                    mshet['lock'].acquire()
                    mcol = mshet['total_col']
                    mkeys = mshet['keys']
                    for mrow in mkeys:
                        if mshet['xls'].getCellValue(mshet['sheet'], mrow, mcol) == None:
                            nrow = writer['xls'].getTotalRows(writer['sheet'])
                            writer['xls'].updateRowValues(writer['sheet'], nrow, mshet['xls'].getRowValues(mshet['sheet'], mrow))
                            writer['xls'].updateCellValue(writer['sheet'], nrow, mshet['total_col'], 'FROM' + mshet['info'])
                    mshet['lock'].release()
                writer['lock'].release()
        else:
            missions = MissionEx(valid_main, 0, staticMethod.checkFrom, args=(None, valid_data, vdict))
            ManagerEx(missions, block_primary=True, new_thread=False).start()
        obj.run = False
        try:
            ExcelWriter.saveBook()
        except:
            messagebox.showinfo("Error", "异常结果储存失败!")
        xfs = ExcelWriter.getExcelObjects(valid_main)
        for xf in xfs:
            try:
                xf.save()
            except:
                messagebox.showinfo("Error", "对比结果储存失败!" + xf.getBookName())
        vdict["var_information"].set('Mission Over')
        time.sleep(3)
        vdict["var_information"].set('')
        
class confirmWindow():
    def __init__(self, root, stop):
        self.root = Toplevel(root)
        self.stop = stop
        self.main_invalid = ExcelWriter.groups['invalid']
        self.data_invalid = ExcelReader.groups['invalid'] 
        self.__createErrorFrame(self.data_invalid, "无法识别的数据文件：").grid(row=0, column=1, sticky=('nsew'))
        self.__createErrorFrame(self.main_invalid, "无法识别的分类至文件：").grid(row=0, column=0, sticky=('nsew'))
        self.__createConfirmFrame().grid(row=1, column=0, columnspan=2, sticky=('nsew'))
    def __createErrorFrame(self, group, text):
        if len(group):
            frame = ttk.Frame(self.root, padding="5")
            lab00 = ttk.Label(frame, text=text)
            txt10 = Text(frame, state='normal')
            sib11 = ttk.Scrollbar(frame, orient='vertical', command=txt10.yview)
            sib20 = ttk.Scrollbar(frame, orient='horizontal', command=txt10.xview)
            for invalid in group:
                if invalid['xls'].type:
                    txt10.insert('end', invalid['sheet'].title + '<' + invalid['xls'].file + '>\n')
                else:
                    txt10.insert('end', invalid['sheet'].name + '<' + invalid['xls'].file + '>\n')
            lab00.grid(row=0, column=0, sticky=('w'))
            txt10.grid(row=1, column=0, sticky=('w', 'e'))
            sib11.grid(row=1, column=1, sticky=('s', 'n', 'w'))
            sib20.grid(row=2, column=0, sticky=('w', 'n', 'e'))
        else:
            frame = ttk.Frame(self.root)
        return frame
    def __createConfirmFrame(self):
        frame = ttk.Frame(self.root, padding="25")
        btn00 = ttk.Button(frame, text="返回修改", padding='15')
        btn01 = ttk.Button(frame, text="继续更新", padding='15')
        btn00.grid(row=0, column=0, sticky=('nsew'))
        btn01.grid(row=0, column=1, sticky=('nsew'))
        btn00.bind('<1>', lambda e: self.quit())
        btn01.bind('<1>', lambda e: self.okay())
        return frame
    def quit(self):
        self.stop.set(True)
        self.root.destroy()
    def okay(self):
        self.stop.set(False)
        self.root.destroy()
        
class tkinterProcdure():
    def __init__(self, app):
        self.app = app
        self.run = False
        self.load = 0
        self.mode = 5
        self.workdir = "E:\\Felix\\Desktop\\test" #os.getcwd()
        self.main_dict = {'name':'main', 'brother':None, 'types':(("Excel Files", "*.xlsx;*.xlsm"), ), 'regular':r"^(?!~).*\.(xlsx|xlsm)$", 'xls':None, 'files':None, 'read_only':False,'listbox_splitter':(), 'listbox_keys':(), 'listbox_copy':()}
        self.data_dict = {'name':'data', 'brother':None, 'types':(("Excel Files", "*.xls;*.xlsx;*.xlsm"), ), 'regular':r"^(?!~).*\.(xls|xlsx|xlsm)$", 'xls':None, 'files':None, 'read_only':True, 'listbox_splitter':(), 'listbox_keys':(), 'listbox_copy':()}
        self.main_dict['brother'], self.data_dict['brother'] = self.data_dict, self.main_dict
    def __getDict(self, dname):
        if dname == 'main':
            return self.main_dict
        elif dname == 'data':
            return self.data_dict
        else:
            raise
    def changeFile(self, vdict):
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        open_files, mdict = [], self.__getDict(vdict['name'])
        if vdict['name'] == 'main' and self.mode < 3:
            return
        elif vdict['name'] == 'main' and self.mode < 5:
            open_file = filedialog.askopenfilename(title="选择模板", initialdir=self.workdir, filetypes=mdict['types'])
            if len(open_file):
                open_files.append(open_file)
        else:
            if messagebox.askyesno("选择", "直接指定文件吗？"):
                open_files = filedialog.askopenfilenames(title="选择文件", initialdir=self.workdir, filetypes=mdict['types'])
            else:
                open_dir = filedialog.askdirectory(title="选择文件夹", initialdir = self.workdir)
                if len(open_dir):
                    open_files = createTreeAsPath(open_dir, fileRegular=mdict['regular'], scanSubFolder=vdict['var_scansub'].get(), treeMode=False)
        if len(open_files):
            bfiles = mdict['brother']['files']
            if bfiles != None:
                if len(set(open_files) & set(bfiles)):
                    messagebox.showinfo("Error", "同一文件不能既是主文件，也是数据文件！")
                    return
            lst, mdict['files'] = vdict['listbox_file'], open_files
            self.workdir = os.path.split(open_files[0])[0]
            staticMethod.resetListBox(lst, open_files, 20)
            lst.selection_set(0)
            self.app.foot_dict["var_information"].set('Loading: ' + open_files[0])
            if mdict['name'] == 'main':
                ExcelWriter.resetGroup('writer')
                mdict['xls'] = staticMethod.loadFile((open_files[0], 'writer', False, self.app.foot_dict, []))  #fname, group, read_only, failed_to = args
            else:
                ExcelReader.resetGroup('reader')
                mdict['xls'] = staticMethod.loadFile((open_files[0], 'reader', True, self.app.foot_dict, []))
            if mdict['xls'] == None:
                self.app.foot_dict["var_information"].set('Fail to load: ' + open_files[0])
                return
            self.app.foot_dict["var_information"].set('Succeed to load: ' + open_files[0])
            self.fileChanged(vdict, mdict)
            if len(open_files) > 1:
                self.load_error = []
                """fname, group, read_only, info_fict, failed_to = args"""
                if mdict['name'] == 'main':
                    print(open_files)
                    missions = MissionEx(open_files[1:], 0, staticMethod.loadFile, (None, 'writer', False, self.app.foot_dict, self.load_error))
                else:
                    print(open_files)
                    missions = MissionEx(open_files[1:], 0, staticMethod.loadFile, (None, 'reader', True, self.app.foot_dict, self.load_error))
                self.load += 1
                ManagerEx(missions, block_primary=True, new_thread=True, end_call=self.backMissionOver).start()
            else:
                self.app.foot_dict["var_information"].set('')
    def fileChanged(self, vdict, mdict):
            xls = mdict['xls']
            sheets = xls.valid_sheets
            for i, sheet in enumerate(sheets):
                try:
                    tlrow = xls.guessTitleRow(sheet, allow_jump=vdict['var_coljump'].get(), not_merge=True, must_unique=True)
                except:
                    continue
                else:
                    staticMethod.resetCombobox(vdict['combobox_sheet'], xls.valid_sheet_names, i)
                    staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), tlrow)
                    staticMethod.resetListBoxes(xls, sheet, tlrow, [vdict['listbox_keys'], vdict['listbox_compare']])
                    staticMethod.autoSelectListboxes(vdict)
                    staticMethod.saveSelectedItems(mdict, vdict)
                    break
            else:
                staticMethod.resetCombobox(vdict['combobox_sheet'], xls.valid_sheet_names, 0)
                staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), 0)
                staticMethod.clearListBoxes([vdict['listbox_keys'], vdict['listbox_compare']])
                staticMethod.autoSelectListboxes(vdict)
                staticMethod.saveSelectedItems(mdict, vdict)
                messagebox.showinfo("Error", "无法识别标题行，请自行设置!")
    def sheetChanged(self, vdict):
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        mdict = self.__getDict(vdict['name'])
        xls = mdict['xls']
        sheet = xls.valid_sheets[vdict['combobox_sheet'].current()]
        try:
            tlrow = xls.guessTitleRow(sheet, allow_jump=vdict['var_coljump'].get(), not_merge=True, must_unique=True)
        except:
            staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), 0)
            staticMethod.clearListBoxes([vdict['listbox_keys'], vdict['listbox_compare']])
            staticMethod.autoSelectListboxes(vdict)
            staticMethod.saveSelectedItems(mdict, vdict)
            messagebox.showinfo("Error", "无法识别标题行，请自行设置!")
        else:
            staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), tlrow)
            staticMethod.resetListBoxes(xls, sheet, tlrow, [vdict['listbox_keys'], vdict['listbox_compare']])
            staticMethod.autoSelectListboxes(vdict)
            staticMethod.saveSelectedItems(mdict, vdict)
    def rowChanged(self, vdict): 
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        mdict = self.__getDict(vdict['name'])
        staticMethod.resetListBoxes(mdict['xls'], mdict['xls'].valid_sheets[vdict['combobox_sheet'].current()], vdict['combobox_row'].current(), [vdict['listbox_keys'], vdict['listbox_compare']])
        staticMethod.autoSelectListboxes(vdict)
        staticMethod.saveSelectedItems(mdict, vdict)
    def selectListboxFile(self, vdict):
        box = vdict['listbox_file']
        if not box.selection_includes(0):
            selected = box.curselection()
            for i in selected:
                box.selection_clear(i)
            box.selection_set(0)
    def selectColumn(self, vdict, bname):
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        mdict = self.__getDict(vdict['name'])
        oldselected = set(mdict[bname])
        newselected = set(vdict[bname].curselection())
        brother = vdict['brother']
        if oldselected > newselected:
            selected = oldselected - newselected
            for i in selected:
                staticMethod.selectListBoxItemByValue(brother[bname], vdict[bname].get(i), False)
        else:
            selected = newselected - oldselected
            for i in selected:
                brselected = staticMethod.selectListBoxItemByValue(brother[bname], vdict[bname].get(i), True)
                if bname in ['listbox_keys', 'listbox_compare']:
                    for name in ['listbox_keys', 'listbox_compare']:
                        if name != bname:
                            if vdict[name].selection_includes(i):
                                vdict[name].selection_clear(i)
                            if brselected > -1:
                                if brother[name].selection_includes(brselected):
                                    brother[name].selection_clear(brselected)
        staticMethod.saveSelectedItems(mdict, vdict)
    def confirmCommand(self, vdict):
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        if self.load:
            messagebox.showinfo("Error", "文件加载中，请稍后再试!")
            return
        if not self.app.checkSelection():
            return
        thrd = threading.Thread(target=staticMethod.confirmCommand, args=(self, vdict))
        thrd.setDaemon(True)
        thrd.start()
    def backMissionOver(self, success, msg, args):
        if self.load:
            self.load -= 1
            info = 'Files loading over!'
        elif self.run:
            self.run = False
            info = 'Data updating over!'
        self.app.foot_dict["var_information"].set(info)
        time.sleep(5)
        self.app.foot_dict["var_information"].set('')
        
class thinterWindow():
    def __init__(self):
        self.root = Tk()
        self.proc = tkinterProcdure(self)
        self.main_dict = {'name':'main'}
        self.data_dict = {'name':'data'}
        self.foot_dict = {'name':'foot'}
        self.body_list = [self.main_dict, self.data_dict]
        self.main_dict['brother'], self.data_dict['brother'] = self.data_dict, self.main_dict
        self.__createLeftFrame(self.main_dict, "验证数据：").grid(row=0, column=0, sticky=('nsew'))
        self.__createRightFrame(self.main_dict).grid(row=0, column=1, columnspan=2, sticky=('nsew'))
        self.__createLeftFrame(self.data_dict, "验证依据：").grid(row=1, column=0, sticky=('nsew'))
        self.__createRightFrame(self.data_dict).grid(row=1, column=1, columnspan=2, sticky=('nsew'))
        self.__createFootFrame(self.foot_dict).grid(row=2, column=0, columnspan=2, sticky=('nsew'))
        self.__createCommandFrame(self.foot_dict).grid(row=2, column=2, sticky=('es'))
        self.__configure()
    def __configure(self):
        self.root.title("Excel数据比对")
        self.root.geometry(self.root.geometry("0x0+200+5"))
        self.main_dict['label_select'].bind('<1>', lambda e: self.proc.changeFile(self.main_dict))
        self.main_dict['listbox_file'].bind('<<ListboxSelect>>', lambda e: self.proc.selectListboxFile(self.main_dict))
        self.main_dict['combobox_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.main_dict))
        self.main_dict['combobox_row'].bind('<<ComboboxSelected>>', lambda e: self.proc.rowChanged(self.main_dict))
        self.main_dict['listbox_keys'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_keys'))
        self.main_dict['listbox_compare'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_compare'))
        
        self.data_dict['label_select'].bind('<1>', lambda e: self.proc.changeFile(self.data_dict))
        self.data_dict['listbox_file'].bind('<<ListboxSelect>>', lambda e: self.proc.selectListboxFile(self.data_dict))
        self.data_dict['combobox_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.data_dict))
        self.data_dict['combobox_row'].bind('<<ComboboxSelected>>', lambda e: self.proc.rowChanged(self.data_dict))
        self.data_dict['listbox_keys'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_keys'))
        self.data_dict['listbox_compare'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_compare'))
        
        self.foot_dict['button_confirm'].bind('<1>', lambda e: self.proc.confirmCommand(self.foot_dict))
    def __createLeftFrame(self, vdict, title):
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text=title)
        lst10 = Listbox(frame, exportselection=False)
        sib12 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        lab30 = ttk.Label(frame, text="表名：", width=8)
        cbx31 = ttk.Combobox(frame, width=8, state="readonly") #state="disabled"
        lab40 = ttk.Label(frame, text="标题行：", width=8)
        cbx41 = ttk.Combobox(frame, width=8, state="readonly") #state="disabled"
        scansub, coljump, onlyone = BooleanVar(value=True), BooleanVar(value=True), BooleanVar(value=False)
        cbn50 = ttk.Checkbutton(frame, variable=scansub, onvalue=True, offvalue=False, text="扫描子目录")
        cbn60 = ttk.Checkbutton(frame, variable=coljump, onvalue=True, offvalue=False, text="标题行允许跳列")
        cbn70 = ttk.Checkbutton(frame, variable=onlyone, onvalue=True, offvalue=False, text="仅当前SHEET")
        lab00.grid(row=0, column=0, sticky=('w'))
        lst10.grid(row=1, column=0, columnspan=2, sticky=('w', 'e'))
        sib12.grid(row=1, column=2, sticky=('s', 'n', 'w'))
        sib20.grid(row=2, column=0, columnspan=2, sticky=('w', 'n', 'e'))
        lab30.grid(row=3, column=0, sticky=('e'))
        cbx31.grid(row=3, column=1, sticky=('w'))
        lab40.grid(row=4, column=0, sticky=('e'))
        cbx41.grid(row=4, column=1, sticky=('w'))
        cbn50.grid(row=5, column=0, sticky=('w'))
        cbn60.grid(row=6, column=0, sticky=('w'))
        cbn70.grid(row=7, column=0, sticky=('w'))
        vdict['label_select'], vdict['listbox_file'] = lab00, lst10
        vdict['combobox_sheet'], vdict['combobox_row'] = cbx31, cbx41
        vdict['var_scansub'], vdict['var_coljump'], vdict['var_onlyone'] = scansub, coljump, onlyone
        return frame
    def __createRightFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text="标识列：")
        lab02 = ttk.Label(frame, text="对比列：")
        lst10 = Listbox(frame, selectmode='multiple', exportselection=False)
        sib11 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        lst12 = Listbox(frame, selectmode='multiple', exportselection=False)
        sib13 = ttk.Scrollbar(frame, orient='vertical', command=lst12.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        sib22 = ttk.Scrollbar(frame, orient='horizontal', command=lst12.xview)
        lab00.grid(row=0, column=0, sticky=('w'))
        lab02.grid(row=0, column=2, sticky=('w'))
        lst10.grid(row=1, column=0, sticky=('n', 'e', 's'))
        sib11.grid(row=1, column=1, sticky=('n', 'w', 's'))
        lst12.grid(row=1, column=2, sticky=('n', 'e', 's'))
        sib13.grid(row=1, column=3, sticky=('n', 'w', 's'))
        sib20.grid(row=2, column=0, sticky=('nsew'))
        sib22.grid(row=2, column=2, sticky=('nsew'))
        vdict['listbox_keys'], vdict['listbox_compare'] = lst10, lst12
        if vdict['name'] == 'main':
            skipsum, summode, notnull = BooleanVar(value=True), BooleanVar(value=False), BooleanVar(value=True)
            ttk.Checkbutton(frame, variable=skipsum, onvalue=True, offvalue=False, text="自动跳过汇总行", state='disabled').grid(row=3, column=0, sticky=('w'))
            ttk.Checkbutton(frame, variable=summode, onvalue=True, offvalue=False, text="加总后对比", state='disabled').grid(row=3, column=2, sticky=('w'))
            ttk.Checkbutton(frame, variable=notnull, onvalue=True, offvalue=False, text="标识列不允许为空").grid(row=4, column=0, sticky=('w'))
            vdict['var_skipsum'], vdict['var_summode'], vdict['var_notnull'] = skipsum, summode, notnull
        return frame
    def __createFootFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="5")
        allmatch, information = BooleanVar(value=False), StringVar()
        ttk.Checkbutton(frame, variable=allmatch, onvalue=True, offvalue=False, text="一致性检查").grid(row=0, column=0, sticky=('w'))
        ttk.Label(frame, text="注：表格要求标题行不重复，不包含合并单元格，且不超过第九行!").grid(row=1, column=0, sticky=('w'))
        ttk.Label(frame, text=" ", textvariable=information).grid(row=2, column=0, sticky=('w'))
        vdict["var_allmatch"], vdict["var_information"] = allmatch, information
        return frame
    def __createCommandFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="10")
        btn00 = ttk.Button(frame, text="执行")
        btn00.grid(row=0, column=0, sticky=('e', 's', 'n'))
        vdict['button_confirm'] = btn00
        return frame
    def checkSelection(self):
        if len(self.data_dict['listbox_keys'].curselection()) == 0:
            messagebox.showinfo("Error", "标识列必须选择!")
            return False
        if len(self.data_dict['listbox_compare'].curselection()) == 0:
            messagebox.showinfo("Error", "对比必须选择!")
            return False
        if len(self.main_dict['listbox_keys'].curselection()) != len(self.data_dict['listbox_keys'].curselection()):
            messagebox.showinfo("Error", "标识列不匹配!")
            return False
        if len(self.main_dict['listbox_compare'].curselection()) != len(self.data_dict['listbox_compare'].curselection()):
            messagebox.showinfo("Error", "对比列不匹配!")
            return False
        return True
    def showMainWindow(self):
        self.root.mainloop()

if __name__ == "__main__":
    staticMethod.main()
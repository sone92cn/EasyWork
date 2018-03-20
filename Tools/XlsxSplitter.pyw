#encoding "UTF-8"
import os, threading, time
from PathTool import createTreeAsPath
from ThreadTool import MissionEx, ManagerEx
from ExcelTool import ExcelReader, ExcelWriter
from FelixFunc import combineLists, wipeNone, createStringFromVarList
from tkinter import Tk, Toplevel, ttk, Listbox, Text, filedialog, messagebox, BooleanVar, StringVar, IntVar
    
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
    @staticmethod
    def checkTitle(args):
        if len(args) == 9:
            shet, check_list, copy_list, split_len, key_len, guess_row, max_row, cls_obj, info_dict = args
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
                shet['split_cols'] = check_pos[:split_len]
                shet['key_cols'] = check_pos[split_len:key_len]
                if len(copy_list):
                    shet['copy_cols'] = shet['xls'].getRowPosByValues(shet['sheet'], row, copy_list, all_included=False)
                else:
                    shet['copy_cols'] = check_pos[key_len:]
                shet['total_col'] = shet['xls'].getTotalCols(shet['sheet'])
                shet['lock'].release()
            info_dict["var_information"].set('Checked: ' + shet['info'])
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
    def updateTo(args):    #search all data shets and auto append
        if len(args) == 4:
            dshet, only_unique, file_mode, new_model = args
            dkeys = dshet['keys']
            for drow in dkeys:
                splits = dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['split_cols'])
                writer = createStringFromVarList(splits, sep='-', add_quotes=False, skip_none=True)
                if writer == '':
                    writer = 'default'
                writers = ExcelWriter.getWriter(writer, file_mode, new_model, not_copy=['keys'], keys={})
                for mshet in writers:
                    mshet['lock'].acquire()
                    if only_unique:
                        if dkeys[drow] in mshet['keys'].values():
                            mshet['lock'].release()
                            continue
                    n = mshet['xls'].getTotalRows(mshet['sheet'])
                    mshet['xls'].updateRowValuesByPos(mshet['sheet'], n, mshet['key_cols'], dkeys[drow], update_all=False)
                    mshet['xls'].updateRowValuesByPos(mshet['sheet'], n, mshet['split_cols'], splits, update_all=False)
                    mshet['xls'].updateRowValuesByPos(mshet['sheet'], n, mshet['copy_cols'], dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['copy_cols'], allow_error=True), update_all=False)
                    mshet['xls'].updateCellValue(mshet['sheet'], n, mshet['total_col'], 'appended from '+dshet['info'], update_all=False)
                    mshet['keys'][n] = dkeys[drow]
                    mshet['lock'].release()
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
        for lname in ['listbox_splitter', 'listbox_keys', 'listbox_copy']:
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
        mdict['listbox_splitter'] = vdict['listbox_splitter'].curselection()
        mdict['listbox_keys'] = vdict['listbox_keys'].curselection()
        mdict['listbox_copy'] = vdict['listbox_copy'].curselection()
        mbdict['listbox_splitter'] = vbdict['listbox_splitter'].curselection()
        mbdict['listbox_keys'] = vbdict['listbox_keys'].curselection()
        mbdict['listbox_copy'] = vbdict['listbox_copy'].curselection()
    @staticmethod
    def getSelectedListboxItems(box, selected):
        items = []
        for i in selected:
            items.append(box.get(i))
        return items
    @staticmethod
    def getValidateSheetDict(vdict, mdict, idict, mode, partmatch=False):
        split = staticMethod.getSelectedListboxItems(vdict['listbox_splitter'], mdict['listbox_splitter'])
        key = staticMethod.getSelectedListboxItems(vdict['listbox_keys'], mdict['listbox_keys'])
        copy = staticMethod.getSelectedListboxItems(vdict['listbox_copy'], mdict['listbox_copy'])
        if partmatch:
            total = combineLists(split, key)
            optional = copy
        else:
            total = combineLists(split, key, copy)
            optional = []
            
        """shet, check_list, key_len, guess_row, max_row, cls_obj, info_dict = args"""
        clen = len(split)
        if mdict['name'] == 'main':
            if mode < 5:
                shets = [ExcelWriter.getGroupShet('writer', mdict['xls'].getSheetObject(vdict['combobox_sheet'].get()))]
            else:
                shets = ExcelWriter.groups['writer']
            ExcelWriter.resetGroup('invalid')
            missions = MissionEx(shets, 0, staticMethod.checkTitle, args=(None, total, optional, clen, clen+len(key), vdict['combobox_row'].current(), 9, ExcelWriter, idict))
        else:
            if vdict['var_onlyone'].get():
                shets = [ExcelReader.getGroupShet('reader', mdict['xls'].getSheetObject(vdict['combobox_sheet'].get()))]
            else:
                shets = ExcelReader.groups['reader']
            ExcelReader.resetGroup('invalid')
            missions = MissionEx(shets, 0, staticMethod.checkTitle, args=(None, total, optional, clen, clen+len(key), vdict['combobox_row'].current(), 9, ExcelReader, idict))
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
        valid_main = []
        if obj.mode > 2:
            valid_main = staticMethod.getValidateSheetDict(vdict_main, mdict_main, vdict, obj.mode)
        valid_data = staticMethod.getValidateSheetDict(vdict_data, mdict_data, vdict, obj.mode, vdict_data['var_partmatch'].get())
        if len(ExcelReader.groups['invalid']) or len(ExcelWriter.groups['invalid']):
            stop = BooleanVar(value=True)
            conf = confirmWindow(obj.app.root, stop)
            obj.app.root.wait_window(conf.root)
            if stop.get():
                obj.run = False
                time.sleep(3)
                vdict["var_information"].set('')
                return
        ExcelWriter.resetBook()
        ExcelWriter.resetWriters()
        if obj.mode < 3:
            file_mode = obj.mode - 1
            model = ExcelReader.getGroupShet('reader', obj.data_dict['xls'].getSheetObject(obj.app.data_dict['combobox_sheet'].get()))
        elif obj.mode < 5:
            file_mode = obj.mode - 3
            model = ExcelWriter.getGroupShet('writer', obj.main_dict['xls'].getSheetObject(obj.app.main_dict['combobox_sheet'].get()))
        else:
            file_mode = 1
            ExcelWriter.createWriterFromGroup() #Create Writers
            model = ExcelWriter.getGroupShet('writer', obj.main_dict['xls'].getSheetObject(obj.app.main_dict['combobox_sheet'].get()))
        """shet, not_null = args"""
        missions = MissionEx(valid_data, 0, staticMethod.checkKeys, args=(None, vdict_data['var_notnull'].get(), vdict), finished=False)
        if vdict_data['var_uniquekey'].get():
            missions.extendMission(valid_main)
        missions.endMission()
        ManagerEx(missions, block_primary=True, new_thread=False).start()
        
        """dshet, only_unique, file_mode = args"""
        missions = MissionEx(valid_data, 0, staticMethod.updateTo, args=(None, vdict_data['var_uniquekey'].get(), file_mode, model))
        
        vdict["var_information"].set('Updating ......!')
        ManagerEx(missions, block_primary=True, new_thread=False).start()
        obj.run = False
        vdict["var_information"].set('Updating completed!')
        if file_mode:
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
                            messagebox.showinfo("Error", "储存分类结果失败：" + xf.getBookName())
        else:
            try:
                ExcelWriter.saveBook()
            except:
                messagebox.showinfo("Error", "对比分类储存失败！")
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
                    missions = MissionEx(open_files[1:], 0, staticMethod.loadFile, args=(None, 'writer', False, self.app.foot_dict, self.load_error))
                else:
                    missions = MissionEx(open_files[1:], 0, staticMethod.loadFile, args=(None, 'reader', True, self.app.foot_dict, self.load_error))
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
                    staticMethod.resetListBoxes(xls, sheet, tlrow, [vdict['listbox_splitter'], vdict['listbox_keys'], vdict['listbox_copy']])
                    staticMethod.autoSelectListboxes(vdict)
                    staticMethod.saveSelectedItems(mdict, vdict)
                    break
            else:
                staticMethod.resetCombobox(vdict['combobox_sheet'], xls.valid_sheet_names, 0)
                staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), 0)
                staticMethod.clearListBoxes([vdict['listbox_splitter'], vdict['listbox_keys'], vdict['listbox_copy']])
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
            staticMethod.clearListBoxes([vdict['listbox_splitter'], vdict['listbox_keys'], vdict['listbox_copy']])
            staticMethod.autoSelectListboxes(vdict)
            staticMethod.saveSelectedItems(mdict, vdict)
            messagebox.showinfo("Error", "无法识别标题行，请自行设置!")
        else:
            staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), tlrow)
            staticMethod.resetListBoxes(xls, sheet, tlrow, [vdict['listbox_splitter'], vdict['listbox_keys'], vdict['listbox_copy']])
            staticMethod.autoSelectListboxes(vdict)
            staticMethod.saveSelectedItems(mdict, vdict)
    def rowChanged(self, vdict): 
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        mdict = self.__getDict(vdict['name'])
        staticMethod.resetListBoxes(mdict['xls'], mdict['xls'].valid_sheets[vdict['combobox_sheet'].current()], vdict['combobox_row'].current(), [vdict['listbox_splitter'], vdict['listbox_keys'], vdict['listbox_copy']])
        staticMethod.autoSelectListboxes(vdict)
        staticMethod.saveSelectedItems(mdict, vdict)
    def modeChanged(self, vdict, mode):
        if mode == 1 or mode == 2:
            if self.mode == 3 or self.mode == 4 or self.mode == 5:
                vdict['listbox_file']['state'], vdict['combobox_sheet']['state'], vdict['combobox_row']['state'], vdict['checkbutton_coljump']['state'], vdict['listbox_splitter']['state'], vdict['listbox_keys']['state'], vdict['listbox_copy']['state']  = \
                            'disabled', 'disabled', 'disabled', 'disabled', 'disabled', 'disabled', 'disabled'
        else:
            if self.mode == 1 or self.mode == 2:
                vdict['listbox_file']['state'], vdict['combobox_sheet']['state'], vdict['combobox_row']['state'], vdict['checkbutton_coljump']['state'], vdict['listbox_splitter']['state'], vdict['listbox_keys']['state'], vdict['listbox_copy']['state']  = \
                            'normal', 'readonly', 'readonly', 'normal', 'normal', 'normal', 'normal'
        self.mode = mode
    def selectListboxFile(self, vdict):
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
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
                if bname in ['listbox_keys', 'listbox_copy']:
                    for name in ['listbox_keys', 'listbox_copy']:
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
    def backMissionOver(self):
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
        self.__createLeftFrame(self.data_dict, "数据来源：").grid(row=0, column=0, sticky=('nsew'))
        self.__createRightFrame(self.data_dict).grid(row=0, column=1, columnspan=2, sticky=('nsew'))
        self.__createLeftFrame(self.main_dict, "分类方式：").grid(row=1, column=0, sticky=('nsew'))
        self.__createRightFrame(self.main_dict).grid(row=1, column=1, columnspan=2, sticky=('nsew'))
        self.__createFootFrame(self.foot_dict).grid(row=2, column=0, columnspan=2, sticky=('nsew'))
        self.__createCommandFrame(self.foot_dict).grid(row=2, column=2, sticky=('es'))
        self.__configure()
    def __configure(self):
        self.root.title("Excel数据分类")
        self.root.geometry(self.root.geometry("0x0+200+5"))
        self.proc.modeChanged(self.main_dict, 1)
        self.main_dict['label_select'].bind('<1>', lambda e: self.proc.changeFile(self.main_dict))
        self.main_dict['listbox_file'].bind('<<ListboxSelect>>', lambda e: self.proc.selectListboxFile(self.main_dict))
        self.main_dict['combobox_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.main_dict))
        self.main_dict['combobox_row'].bind('<<ComboboxSelected>>', lambda e: self.proc.rowChanged(self.main_dict))
        self.main_dict['listbox_splitter'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_splitter'))
        self.main_dict['listbox_keys'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_keys'))
        self.main_dict['listbox_copy'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_copy'))
        self.main_dict['radio_mode1'].bind('<1>', lambda e: self.proc.modeChanged(self.main_dict, 1))
        self.main_dict['radio_mode2'].bind('<1>', lambda e: self.proc.modeChanged(self.main_dict, 2))
        self.main_dict['radio_mode3'].bind('<1>', lambda e: self.proc.modeChanged(self.main_dict, 3))
        self.main_dict['radio_mode4'].bind('<1>', lambda e: self.proc.modeChanged(self.main_dict, 4))
        self.main_dict['radio_mode5'].bind('<1>', lambda e: self.proc.modeChanged(self.main_dict, 5))
        
        self.data_dict['label_select'].bind('<1>', lambda e: self.proc.changeFile(self.data_dict))
        self.data_dict['listbox_file'].bind('<<ListboxSelect>>', lambda e: self.proc.selectListboxFile(self.data_dict))
        self.data_dict['combobox_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.data_dict))
        self.data_dict['combobox_row'].bind('<<ComboboxSelected>>', lambda e: self.proc.rowChanged(self.data_dict))
        self.data_dict['listbox_splitter'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_splitter'))
        self.data_dict['listbox_keys'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_keys'))
        self.data_dict['listbox_copy'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_copy'))
        
        self.foot_dict['button_confirm'].bind('<1>', lambda e: self.proc.confirmCommand(self.foot_dict))
    def __createLeftFrame(self, vdict, title):
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text=title)
        if vdict['name'] == 'main':
            i, mode = 3, IntVar(value=1)
            rbn10 = ttk.Radiobutton(frame, variable=mode, text='自动到SHEET', value = 1)
            rbn11 = ttk.Radiobutton(frame, variable=mode, text='自动到文件', value = 2)
            rbn20 = ttk.Radiobutton(frame, variable=mode, text='模板到SHEET', value = 3)
            rbn21 = ttk.Radiobutton(frame, variable=mode, text='模板到文件', value = 4)
            rbn30 = ttk.Radiobutton(frame, variable=mode, text='分类到已存在的文件', value = 5)
            rbn10.grid(row=1, column=0, sticky=('w'))
            rbn11.grid(row=1, column=1, sticky=('w'))
            rbn20.grid(row=2, column=0, sticky=('w'))
            rbn21.grid(row=2, column=1, sticky=('w'))
            rbn30.grid(row=3, column=0, sticky=('w'))
            vdict['var_mode'] = mode
            vdict['radio_mode1'], vdict['radio_mode2'], vdict['radio_mode3'], vdict['radio_mode4'], vdict['radio_mode5'] = rbn10, rbn11, rbn20, rbn21, rbn30
            scansub, coljump, state = BooleanVar(value=False), BooleanVar(value=True), 'disabled'
        else:
            i, onlyone = 0, BooleanVar(value=False)
            ttk.Checkbutton(frame, variable=onlyone, onvalue=True, offvalue=False, text="仅检索当前SHEET。").grid(row=i+7, column=0, sticky=('w'))
            vdict['var_onlyone'] = onlyone
            scansub, coljump, state = BooleanVar(value=True), BooleanVar(value=True), 'normal'
        lst10 = Listbox(frame, exportselection=False)
        sib12 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        lab30 = ttk.Label(frame, text="表名：", width=8)
        cbx31 = ttk.Combobox(frame, width=8, state="readonly") #state="disabled"
        lab40 = ttk.Label(frame, text="标题行：", width=8)
        cbx41 = ttk.Combobox(frame, width=8, state="readonly") #state="disabled"
        cbn50 = ttk.Checkbutton(frame, variable=scansub, onvalue=True, offvalue=False, text="扫描子目录", state=state)
        cbn60 = ttk.Checkbutton(frame, variable=coljump, onvalue=True, offvalue=False, text="标题行允许跳列")
        lab00.grid(row=0, column=0, sticky=('w'))
        lst10.grid(row=i+1, column=0, columnspan=2, sticky=('w', 'e'))
        sib12.grid(row=i+1, column=2, sticky=('s', 'n', 'w'))
        sib20.grid(row=i+2, column=0, columnspan=2, sticky=('w', 'n', 'e'))
        lab30.grid(row=i+3, column=0, sticky=('e'))
        cbx31.grid(row=i+3, column=1, sticky=('w'))
        lab40.grid(row=i+4, column=0, sticky=('e'))
        cbx41.grid(row=i+4, column=1, sticky=('w'))
        cbn50.grid(row=i+5, column=0, sticky=('w'))
        cbn60.grid(row=i+6, column=0, sticky=('w'))
        vdict['label_select'], vdict['listbox_file'] = lab00, lst10
        vdict['combobox_sheet'], vdict['combobox_row'] = cbx31, cbx41
        vdict['checkbutton_scansub'], vdict['checkbutton_coljump'] = cbn50, cbn60
        vdict['var_scansub'], vdict['var_coljump'] = scansub, coljump
        return frame
    def __createRightFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text="分类依据：")
        lab02 = ttk.Label(frame, text="标识列：")
        lab04 = ttk.Label(frame, text="复制列：")
        lst10 = Listbox(frame, selectmode='multiple', exportselection=False)
        sib11 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        lst12 = Listbox(frame, selectmode='multiple', exportselection=False)
        sib13 = ttk.Scrollbar(frame, orient='vertical', command=lst12.yview)
        lst14 = Listbox(frame, selectmode='multiple', exportselection=False)
        sib15 = ttk.Scrollbar(frame, orient='vertical', command=lst14.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        sib22 = ttk.Scrollbar(frame, orient='horizontal', command=lst12.xview)
        sib24 = ttk.Scrollbar(frame, orient='horizontal', command=lst14.xview)
        lab00.grid(row=0, column=0, sticky=('w'))
        lab02.grid(row=0, column=2, sticky=('w'))
        lab04.grid(row=0, column=4, sticky=('w'))
        lst10.grid(row=1, column=0, sticky=('n', 'e', 's'))
        sib11.grid(row=1, column=1, sticky=('n', 'w', 's'))
        lst12.grid(row=1, column=2, sticky=('n', 'e', 's'))
        sib13.grid(row=1, column=3, sticky=('n', 'w', 's'))
        lst14.grid(row=1, column=4, sticky=('n', 'e', 's'))
        sib15.grid(row=1, column=5, sticky=('n', 'w', 's'))
        sib20.grid(row=2, column=0, sticky=('nsew'))
        sib22.grid(row=2, column=2, sticky=('nsew'))
        sib24.grid(row=2, column=4, sticky=('nsew'))
        vdict['listbox_splitter'], vdict['listbox_keys'], vdict['listbox_copy'] = lst10, lst12, lst14
        if vdict['name'] == 'data':
            jumpsum = BooleanVar(value=True)
            uniquekey = BooleanVar(value=False)
            notnull = BooleanVar(value=True)
            partmatch = BooleanVar(value=False)
            cbn30 = ttk.Checkbutton(frame, variable=jumpsum, onvalue=True, offvalue=False, text="自动跳过汇总行", state='disabled')
            cbn32 = ttk.Checkbutton(frame, variable=uniquekey, onvalue=True, offvalue=False, text="不重复分类")
            cbn34 = ttk.Checkbutton(frame, variable=partmatch, onvalue=True, offvalue=False, text="复制列允许部分匹配")
            cbn42 = ttk.Checkbutton(frame, variable=notnull, onvalue=True, offvalue=False, text="标识列不允许为空")
            cbn30.grid(row=3, column=0, sticky=('w'))
            cbn32.grid(row=3, column=2, sticky=('w'))
            cbn34.grid(row=3, column=4, sticky=('w'))
            cbn42.grid(row=4, column=2, sticky=('w'))
            vdict['var_jumpsum'], vdict['var_uniquekey'], vdict['var_notnull'], vdict['var_partmatch']= jumpsum, uniquekey, notnull, partmatch
        return frame
    def __createFootFrame(self, vdict):
        information = StringVar()
        frame = ttk.Frame(self.root, padding="5")
        ttk.Label(frame, text="注：表格要求标题行不重复，不包含合并单元格，且不超过第九行!").grid(row=0, column=0, sticky=('w'))
        ttk.Label(frame, text=" ", textvariable=information).grid(row=1, column=0, sticky=('w'))
        vdict["var_information"] = information
        return frame
    def __createCommandFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="10")
        btn00 = ttk.Button(frame, text="执行")
        btn00.grid(row=0, column=0, sticky=('e', 's', 'n'))
        vdict['button_confirm'] = btn00
        return frame
    def checkSelection(self):
        if len(self.data_dict['listbox_splitter'].curselection()) == 0:
            messagebox.showinfo("Error", "分类依据必须选择!")
            return False
        
        if len(self.data_dict['listbox_keys'].curselection()) == 0:
            messagebox.showinfo("Error", "标识列必须选择!")
            return False
        if self.proc.mode > 2:
            if len(self.main_dict['listbox_splitter'].curselection()) != len(self.data_dict['listbox_splitter'].curselection()):
                messagebox.showinfo("Error", "分类依据不匹配!")
                return False
            if len(self.main_dict['listbox_keys'].curselection()) != len(self.data_dict['listbox_keys'].curselection()):
                messagebox.showinfo("Error", "标识列不匹配!")
                return False
            if len(self.main_dict['listbox_copy'].curselection()) != len(self.data_dict['listbox_copy'].curselection()):
                messagebox.showinfo("Error", "复制列不匹配!")
                return False
        return True
    def showMainWindow(self):
        self.root.mainloop()

if __name__ == "__main__":
    staticMethod.main()
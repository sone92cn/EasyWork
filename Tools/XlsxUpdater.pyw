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
    @staticmethod
    def checkTitle(args):
        if len(args) == 9:
            shet, check_list, copy_list, key_len, update_len, guess_row, max_row, cls_obj, info_dict = args
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
                shet['update_cols'] = check_pos[key_len:update_len]
                if len(copy_list):
                    shet['copy_cols'] = shet['xls'].getRowPosByValues(shet['sheet'], row, copy_list, all_included=False)
                else:
                    shet['copy_cols'] = check_pos[update_len:]
                shet['update_none'] = [None] * (update_len - key_len)
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
        if len(args) == 3:
            mshets, dshet, null_copy = args
            dkeys = dshet['keys']
            for drow in dkeys:
                values = dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['update_cols'])
                for mshet in mshets:
                    mshet['lock'].acquire()
                    mkeys = mshet['keys']
                    for mrow in mkeys:
                        if mkeys[mrow] == dkeys[drow]:
                            if null_copy:
                                if mshet['xls'].getRowValuesByPos(mshet['sheet'], mrow, mshet['update_cols']) != mshet['update_none']:
                                    break
                            mshet['xls'].updateRowValuesByPos(mshet['sheet'], mrow, mshet['update_cols'], values, update_all=False)
                            mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], 'updated from '+dshet['info'], update_all=False)
                            break
                    else:
                        n = mshet['xls'].getTotalRows(mshet['sheet'])
                        mshet['xls'].updateRowValuesByPos(mshet['sheet'], n, mshet['key_cols'], dkeys[drow], update_all=False)
                        mshet['xls'].updateRowValuesByPos(mshet['sheet'], n, mshet['update_cols'], values, update_all=False)
                        mshet['xls'].updateRowValuesByPos(mshet['sheet'], n, mshet['copy_cols'], dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['copy_cols'], allow_error=True), update_all=False)
                        mshet['xls'].updateCellValue(mshet['sheet'], n, mshet['total_col'], 'appended from '+dshet['info'], update_all=False)
                        mkeys[n] = dkeys[drow]
                    mshet['lock'].release()
        else:
            raise Exception('Args error!')
    @staticmethod
    def updateFrom(args):   #search all main shets not append
        if len(args) == 3:
            mshet, dshets, null_copy = args
            mkeys = mshet['keys']
            for mrow in mkeys:
                if null_copy:
                    if mshet['xls'].getRowValuesByPos(mshet['sheet'], mrow, mshet['update_cols']) != mshet['update_none']:
                        continue
                for dshet in dshets:
                    dkeys = dshet['keys']
                    for drow in dkeys:
                        if mkeys[mrow] == dkeys[drow]:
                            values = dshet['xls'].getRowValuesByPos(dshet['sheet'], drow, dshet['update_cols'])
                            if values == dshet['update_none']:
                                mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], 'found null in '+dshet['info'], update_all=False)
                                continue
                            mshet['lock'].acquire()
                            mshet['xls'].updateRowValuesByPos(mshet['sheet'], mrow, mshet['update_cols'], values, update_all=False)
                            mshet['xls'].updateCellValue(mshet['sheet'], mrow, mshet['total_col'], 'updated from '+dshet['info'], update_all=False)
                            mshet['lock'].release()
                            break
                    else:
                        continue
                    break
        else:
            raise Exception('Args error!')
    @staticmethod
    def clearListBoxes(boxes):
        for box in boxes:
            box.delete(0, 'end')
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
        for lname in ['listbox_check', 'listbox_update', 'listbox_copy']:
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
        mdict['listbox_check'] = vdict['listbox_check'].curselection()
        mdict['listbox_update'] = vdict['listbox_update'].curselection()
        mdict['listbox_copy'] = vdict['listbox_copy'].curselection()
        mbdict['listbox_check'] = vbdict['listbox_check'].curselection()
        mbdict['listbox_update'] = vbdict['listbox_update'].curselection()
        mbdict['listbox_copy'] = vbdict['listbox_copy'].curselection()
    @staticmethod
    def getSelectedListboxItems(box, selected):
        items = []
        for i in selected:
            items.append(box.get(i))
        return items
    @staticmethod
    def getValidateSheetDict(vdict, mdict, idict, partmatch=False):
        check = staticMethod.getSelectedListboxItems(vdict['listbox_check'], mdict['listbox_check'])
        update = staticMethod.getSelectedListboxItems(vdict['listbox_update'], mdict['listbox_update'])
        copy = staticMethod.getSelectedListboxItems(vdict['listbox_copy'], mdict['listbox_copy'])
        if partmatch:
            total = combineLists(check, update)
            optional = copy
        else:
            total = combineLists(check, update, copy)
            optional = []
            
        """shet, check_list, key_len, guess_row, max_row, cls_obj, info_dict = args"""
        clen = len(check)
        if mdict['name'] == 'main':
            if vdict['var_onlyone'].get():
                shets = [ExcelWriter.getGroupShet('writer', mdict['xls'].getSheetObject(vdict['combobox_sheet'].get()))]
            else:
                shets = ExcelWriter.groups['writer']
            ExcelWriter.resetGroup('invalid')
            missions = MissionEx(shets, 0, staticMethod.checkTitle, (None, total, optional, clen, clen+len(update), vdict['combobox_row'].current(), 9, ExcelWriter, idict))
        else:
            if vdict['var_onlyone'].get():
                shets = [ExcelReader.getGroupShet('reader', mdict['xls'].getSheetObject(vdict['combobox_sheet'].get()))]
            else:
                shets = ExcelReader.groups['reader']
            ExcelReader.resetGroup('invalid')
            missions = MissionEx(shets, 0, staticMethod.checkTitle, (None, total, optional, clen, clen+len(update), vdict['combobox_row'].current(), 9, ExcelReader, idict))
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
        valid_data = staticMethod.getValidateSheetDict(vdict_data, mdict_data, vdict, vdict_main['var_partmatch'].get())
        if len(ExcelReader.groups['invalid']) or len(ExcelWriter.groups['invalid']):
            stop = BooleanVar(value=True)
            conf = confirmWindow(obj.app.root, stop)
            obj.app.root.wait_window(conf.root)
            if stop.get():
                obj.run = False
                time.sleep(3)
                vdict["var_information"].set('')
                return
        """shet, not_null = args"""
        missions = MissionEx(valid_main, 0, staticMethod.checkKeys, (None, vdict_main['var_notnull'].get(), vdict), finished=False)
        missions.extendMission(valid_data)
        missions.endMission()
        ManagerEx(missions, block_primary=True, new_thread=False).start()
        if vdict['var_autoappend'].get():
            """mshets, dshet, null_copy = args"""
            missions = MissionEx(valid_data, 1, staticMethod.updateTo, (valid_main, None, vdict_main['var_nullcopy'].get()))
        else:
            """mshet, dshets, null_copy = args"""
            missions = MissionEx(valid_main, 0, staticMethod.updateFrom, (None, valid_data, vdict_main['var_nullcopy'].get()))
        vdict["var_information"].set('Updating ......!')
        ManagerEx(missions, block_primary=True, new_thread=False).start()
        obj.run = False
        vdict["var_information"].set('Updating completed!')
        xfs = ExcelWriter.getExcelObjects(valid_main)
        for xf in xfs:
            try:
                xf.save()
            except:
                messagebox.showinfo("Error", "更新结果储存失败!" + xf.getBookName())
        vdict["var_information"].set('Mission Over')
        time.sleep(3)
        vdict["var_information"].set('')
        
class confirmWindow():
    def __init__(self, root, stop):
        self.root = Toplevel(root)
        self.stop = stop
        self.main_invalid = ExcelWriter.groups['invalid']
        self.data_invalid = ExcelReader.groups['invalid']
        self.__createErrorFrame(self.main_invalid, "无法识别的主文件：").grid(row=0, column=0, sticky=('nsew'))
        self.__createErrorFrame(self.data_invalid, "无法识别的数据文件：").grid(row=0, column=1, sticky=('nsew'))
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
        self.workdir = "E:\\Felix\\Desktop\\test" #os.getcwd()
        self.main_dict = {'name':'main', 'brother':None, 'types':(("Excel Files", "*.xlsx;*.xlsm"), ), 'regular':r"^(?!~).*\.(xlsx|xlsm)$", 'xls':None, 'files':None, 'read_only':False,'listbox_check':(), 'listbox_update':(), 'listbox_copy':()}
        self.data_dict = {'name':'data', 'brother':None, 'types':(("Excel Files", "*.xls;*.xlsx;*.xlsm"), ), 'regular':r"^(?!~).*\.(xls|xlsx|xlsm)$", 'xls':None, 'files':None, 'read_only':True, 'listbox_check':(), 'listbox_update':(), 'listbox_copy':()}
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
        open_files = []
        lst = vdict['listbox_file']
        mdict = self.__getDict(vdict['name'])
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
            mdict['files'] = open_files
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
                    missions = MissionEx(open_files[1:], 0, staticMethod.loadFile, (None, 'writer', False, self.app.foot_dict, self.load_error))
                else:
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
                    staticMethod.resetListBoxes(xls, sheet, tlrow, [vdict['listbox_check'], vdict['listbox_update'], vdict['listbox_copy']])
                    staticMethod.autoSelectListboxes(vdict)
                    staticMethod.saveSelectedItems(mdict, vdict)
                    break
            else:
                staticMethod.resetCombobox(vdict['combobox_sheet'], xls.valid_sheet_names, 0)
                staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), 0)
                staticMethod.clearListBoxes([vdict['listbox_check'], vdict['listbox_update'], vdict['listbox_copy']])
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
            staticMethod.clearListBoxes([vdict['listbox_check'], vdict['listbox_update'], vdict['listbox_copy']])
            staticMethod.autoSelectListboxes(vdict)
            staticMethod.saveSelectedItems(mdict, vdict)
            messagebox.showinfo("Error", "无法识别标题行，请自行设置!")
        else:
            staticMethod.resetRowCombobox(vdict['combobox_row'], min(xls.getTotalRows(sheet), 8), tlrow)
            staticMethod.resetListBoxes(xls, sheet, tlrow, [vdict['listbox_check'], vdict['listbox_update'], vdict['listbox_copy']])
            staticMethod.autoSelectListboxes(vdict)
            staticMethod.saveSelectedItems(mdict, vdict)
        
    def rowChanged(self, vdict): 
        if self.run:
            messagebox.showinfo("Error", "后台程序正在运行，请稍后再试!")
            return
        mdict = self.__getDict(vdict['name'])
        staticMethod.resetListBoxes(mdict['xls'], mdict['xls'].valid_sheets[vdict['combobox_sheet'].current()], vdict['combobox_row'].current(), [vdict['listbox_check'], vdict['listbox_update'], vdict['listbox_copy']])
        staticMethod.autoSelectListboxes(vdict)
        staticMethod.saveSelectedItems(mdict, vdict)
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
                for name in ['listbox_check', 'listbox_update', 'listbox_copy']:
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
        self.__createLeftFrame(self.main_dict, "主文件：").grid(row=0, column=0, sticky=('nsew'))
        self.__createRightFrame(self.main_dict).grid(row=0, column=1, columnspan=2, sticky=('nsew'))
        self.__createLeftFrame(self.data_dict, "数据来源：").grid(row=1, column=0, sticky=('nsew'))
        self.__createRightFrame(self.data_dict).grid(row=1, column=1, columnspan=2, sticky=('nsew'))
        self.__createFootFrame(self.foot_dict).grid(row=2, column=0, columnspan=2, sticky=('nsew'))
        self.__createCommandFrame(self.foot_dict).grid(row=2, column=2, sticky=('es'))
        self.__configure()
    def __configure(self):
        self.root.title("Excel数据更新")
        self.root.geometry(self.root.geometry("0x0+200+10"))
        self.main_dict['var_scansub'].set(True)
        self.main_dict['var_coljump'].set(True)
        self.main_dict['var_onlyone'].set(False)
        self.main_dict['var_notnull'].set(True)
        self.main_dict['var_nullcopy'].set(True)
        self.main_dict['var_jumpsum'].set(True)
        self.main_dict['var_partmatch'].set(False)
        self.main_dict['listbox_file'].bind('<<ListboxSelect>>', lambda e: self.proc.selectListboxFile(self.main_dict))
        self.main_dict['reset_file'].bind('<1>', lambda e: self.proc.changeFile(self.main_dict))
        self.main_dict['combobox_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.main_dict))
        self.main_dict['combobox_row'].bind('<<ComboboxSelected>>', lambda e: self.proc.rowChanged(self.main_dict))
        self.main_dict['listbox_check'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_check'))
        self.main_dict['listbox_update'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_update'))
        self.main_dict['listbox_copy'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.main_dict, 'listbox_copy'))
        self.data_dict['var_scansub'].set(True)
        self.data_dict['var_coljump'].set(True)
        self.data_dict['var_onlyone'].set(False)
        self.data_dict['listbox_file'].bind('<<ListboxSelect>>', lambda e: self.proc.selectListboxFile(self.data_dict))
        self.data_dict['reset_file'].bind('<1>', lambda e: self.proc.changeFile(self.data_dict))
        self.data_dict['combobox_sheet'].bind('<<ComboboxSelected>>', lambda e: self.proc.sheetChanged(self.data_dict))
        self.data_dict['combobox_row'].bind('<<ComboboxSelected>>', lambda e: self.proc.rowChanged(self.data_dict))
        self.data_dict['listbox_check'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_check'))
        self.data_dict['listbox_update'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_update'))
        self.data_dict['listbox_copy'].bind('<<ListboxSelect>>', lambda e: self.proc.selectColumn(self.data_dict, 'listbox_copy'))
        self.foot_dict['button_confirm'].bind('<1>', lambda e: self.proc.confirmCommand(self.foot_dict))
        self.foot_dict['var_autoappend'].set(False)
    def __createLeftFrame(self, vdict, title):
        scansub = BooleanVar()
        coljump = BooleanVar()
        onlyone = BooleanVar()
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text=title)
        lst10 = Listbox(frame, exportselection=False)
        sib12 = ttk.Scrollbar(frame, orient='vertical', command=lst10.yview)
        sib20 = ttk.Scrollbar(frame, orient='horizontal', command=lst10.xview)
        lab30 = ttk.Label(frame, text="表名：", width=8)
        cbx31 = ttk.Combobox(frame, width=8, state="readonly") #state="disabled"
        lab40 = ttk.Label(frame, text="标题行：", width=8)
        cbx41 = ttk.Combobox(frame, width=8, state="readonly") #state="disabled"
        cbn50 = ttk.Checkbutton(frame, variable=scansub, onvalue=True, offvalue=False, text="扫描子目录")
        cbn60 = ttk.Checkbutton(frame, variable=coljump, onvalue=True, offvalue=False, text="标题行允许跳列")
        cbn70 = ttk.Checkbutton(frame, variable=onlyone, onvalue=True, offvalue=False, text="仅检索当前SHEET。")
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
        vdict['reset_file'] = lab00
        vdict['listbox_file'] = lst10
        vdict['combobox_sheet'] = cbx31
        vdict['combobox_row'] = cbx41
        vdict['var_scansub'] = scansub
        vdict['var_coljump'] = coljump
        vdict['var_onlyone'] = onlyone
        return frame
    def __createRightFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="5")
        lab00 = ttk.Label(frame, text="标识列：")
        lab02 = ttk.Label(frame, text="更新列：")
        lab04 = ttk.Label(frame, text="追加复制列：")
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
        vdict['listbox_check'] = lst10
        vdict['listbox_update'] = lst12
        vdict['listbox_copy'] = lst14
        if vdict['name'] == 'main':
            notnull = BooleanVar()
            nullcopy = BooleanVar()
            jumpsum = BooleanVar()
            partmatch = BooleanVar()
            cbn30 = ttk.Checkbutton(frame, variable=notnull, onvalue=True, offvalue=False, text="标识列不允许为空")
            cbn32 = ttk.Checkbutton(frame, variable=nullcopy, onvalue=True, offvalue=False, text="被更新列无数据时更新")
            cbn34 = ttk.Checkbutton(frame, variable=partmatch, onvalue=True, offvalue=False, text="追加复制列允许部分匹配")
            cbn40 = ttk.Checkbutton(frame, variable=jumpsum, onvalue=True, offvalue=False, text="自动跳过汇总行", state='disabled')
            cbn30.grid(row=3, column=0, sticky=('w'))
            cbn32.grid(row=3, column=2, sticky=('w'))
            cbn34.grid(row=3, column=4, sticky=('w'))
            cbn40.grid(row=4, column=0, sticky=('w'))
            vdict['var_notnull'] = notnull
            vdict['var_nullcopy'] = nullcopy
            vdict['var_jumpsum'] = jumpsum
            vdict['var_partmatch'] = partmatch
        return frame
    def __createFootFrame(self, vdict):
        autoappend = BooleanVar()
        information = StringVar()
        frame = ttk.Frame(self.root, padding="5")
        cbn00 = ttk.Checkbutton(frame, variable=autoappend, onvalue=True, offvalue=False, text="数据文件包含主表不存在的数据时，自动追加主表数据。")
        lab10 = ttk.Label(frame, text="注：表格要求标题行不重复，不包含合并单元格，且不超过第九行!")
        lab20 = ttk.Label(frame, text=" ", textvariable=information)
        cbn00.grid(row=0, column=0, sticky=('w'))
        lab10.grid(row=1, column=0, sticky=('w'))
        lab20.grid(row=2, column=0, sticky=('w'))
        vdict["var_autoappend"] = autoappend
        vdict["var_information"] = information
        return frame
    def __createCommandFrame(self, vdict):
        frame = ttk.Frame(self.root, padding="25")
        btn00 = ttk.Button(frame, text="更新")
        btn00.grid(row=0, column=0, sticky=('e', 's', 'n'))
        vdict['button_confirm'] = btn00
        return frame
    def checkSelection(self):
        if len(self.main_dict['listbox_check'].curselection()) == 0 or len(self.data_dict['listbox_check'].curselection()) == 0:
            messagebox.showinfo("Error", "标识列必须选择!")
            return False
        if len(self.main_dict['listbox_check'].curselection()) != len(self.data_dict['listbox_check'].curselection()):
            messagebox.showinfo("Error", "主文件与数据文件标识列不匹配!")
            return False
        if not self.foot_dict['var_autoappend'].get():
            if len(self.main_dict['listbox_update'].curselection()) == 0 or len(self.data_dict['listbox_update'].curselection()) == 0:
                messagebox.showinfo("Error", "主文件不自动追加时，更新列必选!")
                return False
        if len(self.main_dict['listbox_update'].curselection()) != len(self.data_dict['listbox_update'].curselection()):
            messagebox.showinfo("Error", "主文件与数据文件更新列不匹配!")
            return False
        if len(self.main_dict['listbox_copy'].curselection()) != len(self.data_dict['listbox_copy'].curselection()):
            messagebox.showinfo("Error", "主文件与数据文件追加复制列不匹配!")
            return False
        if len(self.main_dict['listbox_update'].curselection()) == 0 and len(self.main_dict['listbox_copy'].curselection()) == 0:
            messagebox.showinfo("Error", "文件更新列和追加复制列不能均不选!")
            return False
        return True
    def showMainWindow(self):
        self.root.mainloop()

if __name__ == "__main__":
    staticMethod.main()
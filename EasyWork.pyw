#version=2016.10.29.001
from ctypes import cdll, create_string_buffer, pointer
import os, re, sys, shutil, win32con, win32gui, win32api, win32clipboard
from PathTool import createTreeAsPath, createDirAsPath, createDirAsTree, getExtFromPath, removeEmptyDir

CID_REFR = 1
CID_EXIT = 2

OPT_OPEN = 1
OPT_COPY = 2
OPT_PAST = 3
OPT_SAVE = 4

class staticMethod():
    @staticmethod
    def main():
        if __name__ == "__main__":
            wndclassname, wndtitlename = "MyPythonWin32EasyWorkMainClass", "EasyWork"
            if win32gui.FindWindow(wndclassname, wndtitlename):
                win32api.MessageBox(win32con.NULL, '程序已运行！', '错误', win32con.MB_OK)
                sys.exit()
            os.chdir(os.path.split(os.path.realpath(__file__))[0])
            WindowApplication(wndclassname, wndtitlename).startHandleMessages()
    @staticmethod
    def copyFile(fpath, tpath, move=False):
        if fpath == tpath:
            return True
        if os.path.exists(tpath):
            if 1 == win32gui.MessageBox(win32con.NULL, tpath + "\n目标文件已存在，确定替换吗？", "提示", win32con.MB_OKCANCEL):
                try:
                    os.remove(tpath)
                except:
                    win32gui.MessageBox(win32con.NULL, tpath + "\n目标文件无法删除，操作失败！", "错误", win32con.MB_OK)
                    return False
            else:
                return False
        if move:
            try:
                shutil.move(fpath, tpath)
            except:
                win32gui.MessageBox(win32con.NULL, tpath + "\n移动文件失败！", "错误", win32con.MB_OK)
                return False
        else:
            try:
                shutil.copy(fpath, tpath)
            except:
                win32gui.MessageBox(win32con.NULL, tpath + "\n复制文件失败！", "错误", win32con.MB_OK)
                return False
        return True
    @staticmethod
    def copyFiles(files, tdir, move):
        if not os.path.isdir(tdir):
            win32gui.MessageBox(win32con.NULL, tdir + "\目标文件夹不存在，任务取消！", "错误", win32con.MB_OK)
            return False
        for path in files:
            if os.path.isfile(path):
                staticMethod.copyFile(path, os.path.join(tdir, os.path.split(path)[1]), move)
            elif os.path.isdir(path):
                dname = tdir
                fname = os.path.split(path)[1]
                tree = createTreeAsPath(path, treeMode=True, forFile=False)
                if len(fname):
                    if createDirAsPath(os.path.join(tdir, fname)):
                        dname = os.path.join(tdir, fname)
                if createDirAsTree(dname, tree):
                    indx = len(path) + 1
                    lsts = createTreeAsPath(path)
                    for lst in lsts:
                        staticMethod.copyFile(lst, os.path.join(dname, lst[indx:]), move)
                if move:
                    removeEmptyDir(path)
    @staticmethod
    def getSettings(filepath):
        file = open(filepath, encoding="utf8")
        settings = {"[CMD1]":[], "[CMD2]":[], "[DIRS]":[], "[EXTS]":[], "[SAVS]":[]}
        act = ""
        line = file.readline()
        while len(line):
            line = line.replace("\r", "").replace("\n", "")
            if line in settings.keys():
                act = line  #print(line)
            elif len(act):
                spt = line.split("=")
                if len(spt) == 2:
                    settings[act].append((spt[0], spt[1]))
            elif line == "[ENDS]":
                break
            line = file.readline()
        file.close()
        return settings
    @staticmethod
    def createSubMenu(menu, tree,  parent, cmdid, cmdlist, forFile, igRe, opt=OPT_OPEN):
        if opt == OPT_SAVE:
            win32gui.AppendMenu(menu, win32con.MF_STRING, cmdid, '<当前目录>')
            cmdlist[cmdid] = [opt, parent]
            cmdid += 1
        hasFile = False
        for branch in tree:
            if len(branch[2]):
                submenu = win32gui.CreatePopupMenu()
                cmdid = staticMethod.createSubMenu(submenu, branch[2], os.path.join(parent, branch[1]), cmdid, cmdlist, forFile, igRe)
                win32gui.AppendMenu(menu, win32con.MF_POPUP, submenu, branch[1])
            else:
                if igRe != None:
                    if igRe.match(branch[1]) != None:
                        win32gui.AppendMenu(menu, win32con.MF_STRING, cmdid, os.path.splitext(branch[1])[0])
                    else:
                        win32gui.AppendMenu(menu, win32con.MF_STRING, cmdid, branch[1])
                else:
                    win32gui.AppendMenu(menu, win32con.MF_STRING, cmdid, branch[1])
                cmdlist[cmdid] = [opt, os.path.join(parent, branch[1])]
                cmdid += 1
                hasFile = True
        if forFile:
            if hasFile:
                #win32gui.InsertMenu(menu, 0, win32con.MF_BYPOSITION|win32con.MF_STRING, cmdid, '<拷出资料>') 
                win32gui.AppendMenu(menu, win32con.MF_STRING, cmdid, '<拷出文件>')
                cmdlist[cmdid] = [OPT_COPY, parent]
                cmdid += 1
        elif opt == OPT_OPEN:
            #win32gui.InsertMenu(menu, 0, win32con.MF_BYPOSITION|win32con.MF_STRING, cmdid, '<拷入资料>') 
            win32gui.AppendMenu(menu, win32con.MF_STRING, cmdid, '<粘入文件>')
            cmdlist[cmdid] = [OPT_PAST, parent]
            cmdid += 1
        return cmdid

class WindowFunction():
    def __init__(self, app):
        self.app = app
        self.lmenu = None
        self.rmenu = None
        self.cmdlist = {}
        self.cmdid = 1000
        self.mode = ''
        self.wkdir = ''
        self.dll = cdll.LoadLibrary('libCBrowseDirDll.dll');
    def __createPopupMenu(self, fname, forFile):
        rstr = ""
        istr = ""
        setting = staticMethod.getSettings(fname)
        if forFile:
            for ext in setting["[EXTS]"]:
                e = ext[1]
                if e[0] == ',' or e[0] == '，':
                    e = e[1:]
                    if len(istr):
                        istr = istr + "|" + e
                    else:
                        istr = r'^(?!~).*\.(' + e
                if len(rstr):
                    rstr = rstr + "|" + e
                else:
                    rstr = r'^(?!~).*\.(' + e
            rstr = rstr + ")$"
            self.regul = rstr
        if len(istr):
            istr = istr + ")$"
            igre = re.compile(istr, re.IGNORECASE)
        else:
            igre = None
        menu = win32gui.CreatePopupMenu()   
        for paras in setting["[CMD1]"]:
            win32gui.AppendMenu(menu, win32con.MF_STRING, self.cmdid, paras[0])
            self.cmdlist[self.cmdid] = [OPT_OPEN, paras[1]]
            self.cmdid += 1
        if len(setting["[CMD1]"]):
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "");
            
        for paras in setting["[DIRS]"]:
            submenu = win32gui.CreatePopupMenu()
            if ";" in paras[1]:
                cmds = paras[1].split(";")
                cmid = self.cmdid
                for cmd in cmds:
                    if "," in cmd:
                        cmd, nlp = cmd.split(",")
                        nlp = int(nlp)
                    else:
                        nlp = 100
                    tree = createTreeAsPath(cmd, rstr, True, True, 0, forFile, nlp)
                    if len(tree):
                        self.cmdid = staticMethod.createSubMenu(submenu, tree, cmd, self.cmdid, self.cmdlist, forFile, igre)
                if cmid < self.cmdid:
                    win32gui.AppendMenu(menu, win32con.MF_POPUP, submenu, paras[0])
            else:
                cmd = paras[1]
                if "," in cmd:
                    cmd, nlp = cmd.split(",")
                    nlp = int(nlp)
                else:
                    nlp = 100
                tree = createTreeAsPath(cmd, rstr, True, True, 0, forFile, nlp)
                if len(tree):
                    self.cmdid = staticMethod.createSubMenu(submenu, tree, cmd, self.cmdid, self.cmdlist, forFile, igre)
                    win32gui.AppendMenu(menu, win32con.MF_POPUP, submenu, paras[0])
        if len(setting["[DIRS]"]):
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "");
            
        if not forFile:  
            if len(setting["[SAVS]"]):
                smenu = win32gui.CreatePopupMenu()
                for paras in setting["[SAVS]"]:
                    if ";" in paras[1]:
                        continue
                    else:
                        if "," in paras[1]:
                            cmd, nlp = paras[1].split(",")
                            nlp = int(nlp)
                            tree = createTreeAsPath(cmd, rstr, True, True, 0, forFile, nlp)
                            if len(tree):
                                submenu = win32gui.CreatePopupMenu()
                                self.cmdid = staticMethod.createSubMenu(submenu, tree, cmd, self.cmdid, self.cmdlist, forFile, igre, OPT_SAVE)
                                win32gui.AppendMenu(smenu, win32con.MF_POPUP, submenu, paras[0])
                        else:
                            win32gui.AppendMenu(smenu, win32con.MF_STRING, self.cmdid, paras[0])
                            self.cmdlist[self.cmdid] = [OPT_SAVE, paras[1]]
                            self.cmdid += 1
                win32gui.AppendMenu(menu, win32con.MF_POPUP, smenu, "文件归档")
                win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "");
            
        for paras in setting["[CMD2]"]:
            win32gui.AppendMenu(menu, win32con.MF_STRING, self.cmdid, paras[0])
            self.cmdlist[self.cmdid] = [OPT_OPEN, paras[1]]
            self.cmdid += 1
        if len(setting["[CMD2]"]):
            win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "");
            
        win32gui.AppendMenu(menu, win32con.MF_STRING, 0, "隐藏")
        win32gui.AppendMenu(menu, win32con.MF_STRING, self.cmdid, "设置")
        self.cmdlist[self.cmdid] = [OPT_OPEN, fname]
        self.cmdid += 1
        
        win32gui.AppendMenu(menu, win32con.MF_STRING, CID_REFR, "刷新")
        win32gui.AppendMenu(menu, win32con.MF_STRING, CID_EXIT, "退出")
        
        return menu
    def createTaskBar(self):
        iconPathName = "Easywork.ico"
        if os.path.isfile(iconPathName): # Look in DLLs dir, a-la py 2.5
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(self.hInstance, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        else:
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)  #win32api.MessageBox(win32con.NULL, '图标文件不存在，使用默认图标！', '错误', win32con.MB_OK)
    
        nid = (self.app.hwnd, 0, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP, win32con.WM_USER, hicon, "Designed by Felix")
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        except win32gui.error:
            win32api.MessageBox(win32con.NULL, '添加任务栏图标失败，程序终止！', '错误', win32con.MB_OK)
            win32gui.DestroyWindow(self.app.hwnd)
    def createPopupMenu(self):
        self.lmenu = self.__createPopupMenu(os.path.join(sys.path[0], "settings_f.txt"), True)
        self.rmenu = self.__createPopupMenu(os.path.join(sys.path[0], "settings_d.txt"), False)
    def prepareWindow(self, wndclassname, wndtitlename):
        self.hInstance = win32api.GetModuleHandle(None)
        wc = win32gui.WNDCLASS();
        wc.hInstance = self.hInstance
        wc.lpszClassName = wndclassname;
        wc.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW | win32con.CS_GLOBALCLASS;
        wc.hCursor = win32api.LoadCursor( 0, win32con.IDC_ARROW )
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.lpfnWndProc = self.WindowProcdure # could also specify a wndproc.
        wc.hIcon = win32con.NULL;
        wc.lpszMenuName = "";
        if not(win32gui.RegisterClass(wc)):
            win32api.MessageBox(win32con.NULL, '注册窗口类失败！', '错误', win32con.MB_OK)
            sys.exit()
        
        self.hwnd = win32gui.CreateWindowEx(win32con.WS_EX_TOOLWINDOW, wc.lpszClassName, wndtitlename, \
                                            win32con.WS_OVERLAPPED|win32con.WS_SYSMENU, 300, 200, 320, 225, 0, 0, self.hInstance, None)     #win32con.WS_EX_TOPMOST|
        self.cinfo = win32gui.CreateWindow("Static", "复制文件：", win32con.WS_CHILD, \
                                    20,  20, 200, 25, self.hwnd, self.hInstance, 0, None)
        self.cfrom = win32gui.CreateWindow("ComboBox", "", win32con.WS_CHILD|win32con.CBS_DROPDOWNLIST|win32con.CBS_HASSTRINGS|win32con.WS_VSCROLL, \
                                    100, 20, 180, 100, self.hwnd, self.hInstance, 0, None)
        self.ctext = win32gui.CreateWindow("Static", "重命名为：", win32con.WS_CHILD, \
                                    20,  50, 200, 25, self.hwnd, self.hInstance, 0, None)
        self.cname= win32gui.CreateWindow("Edit", "", win32con.WS_CHILD, \
                                    100, 50, 180, 18, self.hwnd, self.hInstance, 0, None)
        self.ccurf = win32gui.CreateWindow("Button", "到当前文件夹", win32con.WS_CHILD|win32con.BS_AUTORADIOBUTTON, \
                                    20,  75, 250, 25, self.hwnd, self.hInstance, 0, None)
        self.cself = win32gui.CreateWindow("Button", "到指定文件夹", win32con.WS_CHILD|win32con.BS_AUTORADIOBUTTON, \
                                    20, 100, 250, 25, self.hwnd, self.hInstance, 0, None)
        self.cmove = win32gui.CreateWindow("Button", "完成后删除原文件", win32con.WS_CHILD|win32con.BS_AUTOCHECKBOX, \
                                    20, 125, 120, 25, self.hwnd, self.hInstance, 0, None)
        self.cfirm = win32gui.CreateWindow("Button", "确定", win32con.WS_CHILD|win32con.BS_PUSHBUTTON, \
                                    220,145,  60, 25, self.hwnd, self.hInstance, 0, None)
        self.clist = [self.cinfo, self.cfrom, self.ctext, self.cname, self.ccurf, self.cself, self.cmove, self.cfirm]
        
        self.pinfo = win32gui.CreateWindow("Static", "粘贴到：", win32con.WS_CHILD, \
                                    20,  20, 200, 25, self.hwnd, self.hInstance, 0, None)
        self.pcurf = win32gui.CreateWindow("Button", "当前文件夹", win32con.WS_CHILD|win32con.BS_AUTORADIOBUTTON|win32con.WS_GROUP, \
                                    20,  45, 250, 25, self.hwnd, self.hInstance, 0, None)
        self.psubf = win32gui.CreateWindow("Button", "子文件夹", win32con.WS_CHILD|win32con.BS_AUTORADIOBUTTON, \
                                    20,  70, 80, 25, self.hwnd, self.hInstance, 0, None)
        self.psubs = win32gui.CreateWindow("ComboBox", "", win32con.WS_CHILD|win32con.CBS_DROPDOWNLIST|win32con.CBS_HASSTRINGS|win32con.WS_VSCROLL, \
                                    120, 72, 160, 100, self.hwnd, self.hInstance, 0, None)
        self.pnewf = win32gui.CreateWindow("Button", "新建子文件夹", win32con.WS_CHILD|win32con.BS_AUTORADIOBUTTON, \
                                    20,  95, 100, 25, self.hwnd, self.hInstance, 0, None)
        self.pname = win32gui.CreateWindow("Edit", "", win32con.WS_CHILD, \
                                    120, 95, 160, 20, self.hwnd, self.hInstance, 0, None)
        self.pself = win32gui.CreateWindow("Button", "指定的文件夹", win32con.WS_CHILD|win32con.BS_AUTORADIOBUTTON, \
                                    20, 120, 250, 25, self.hwnd, self.hInstance, 0, None)
        self.pmove = win32gui.CreateWindow("Button", "完成后删除原文件", win32con.WS_CHILD|win32con.BS_AUTOCHECKBOX, \
                                    20, 145, 120, 25, self.hwnd, self.hInstance, 0, None)
        self.pfirm = win32gui.CreateWindow("Button", "确定", win32con.WS_CHILD|win32con.BS_PUSHBUTTON, \
                                    220, 145, 60, 25, self.hwnd, self.hInstance, 0, None)
        win32gui.SendMessage(self.pmove, win32con.BM_SETCHECK, 1, 0)
        self.plist = [self.pinfo, self.pcurf, self.psubf, self.psubs, self.pnewf, self.pname, self.pself, self.pmove, self.pfirm]
        
        hfont = win32gui.GetStockObject(17) #DEFAULT_GUI_FONT 17
        for hwnd in self.clist:
            win32gui.SendMessage(hwnd, win32con.WM_SETFONT, hfont, 1)  #设置控件字体 
        for hwnd in self.plist:
            win32gui.SendMessage(hwnd, win32con.WM_SETFONT, hfont, 1)  #设置控件字体 
        return self.hwnd
    def registerHotKey(self):
        pass
    def switchMode(self, mode, path):
        self.wkdir = path
        if mode == 'c':
            if self.mode != mode:
                for hwnd in self.plist:
                    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
                for hwnd in self.clist:
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SendMessage(self.ccurf, win32con.BM_SETCHECK, 1, 0)
            win32gui.SendMessage(self.cself, win32con.BM_SETCHECK, 0, 0)
            win32gui.SetWindowText(self.cname, "")
            win32gui.SetWindowText(self.ccurf, "到当前文件夹（" + os.path.split(path)[1] + "）")
            win32gui.SetWindowText(self.cself, "到指定文件夹")
            lsts = createTreeAsPath(path, self.regul, False, False, True, True)
            win32gui.SendMessage(self.cfrom, win32con.CB_RESETCONTENT, 0, 0)
            for lst in lsts:
                win32gui.SendMessage(self.cfrom, win32con.CB_ADDSTRING, 0, lst)
        elif mode == 'p':
            if self.mode != mode:
                for hwnd in self.clist:
                    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
                for hwnd in self.plist:
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.SendMessage(self.pcurf, win32con.BM_SETCHECK, 1, 0)
            win32gui.SendMessage(self.psubf, win32con.BM_SETCHECK, 0, 0)
            win32gui.SendMessage(self.pnewf, win32con.BM_SETCHECK, 0, 0)
            win32gui.SendMessage(self.pself, win32con.BM_SETCHECK, 0, 0)
            win32gui.EnableWindow(self.psubs, False)
            win32gui.EnableWindow(self.pname, False)
            win32gui.SetWindowText(self.pcurf, "当前文件夹（" + os.path.split(path)[1] + "）")
            win32gui.SetWindowText(self.pname, "")
            lsts = createTreeAsPath(path, "", False, False, len(path), False)
            win32gui.SendMessage(self.psubs, win32con.CB_RESETCONTENT, 0, 0)
            for lst in lsts:
                win32gui.SendMessage(self.psubs, win32con.CB_ADDSTRING, 0, lst)
        self.mode = mode
    def WindowProcdure(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_USER:
            return self.OnTaskbar(hwnd, wparam, lparam)
        elif msg == win32con.WM_COMMAND:
            return self.OnCommand(hwnd, wparam, lparam)
        elif msg == win32con.WM_CLOSE:
            if self.updated:
                self.createPopupMenu()
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        elif msg == win32con.WM_DESTROY:
            nid = (hwnd, 0)
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
            win32gui.PostQuitMessage(0)
        else:
            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    def OnTaskbar(self, hwnd, wparam, lparam):
        if lparam==win32con.WM_LBUTTONUP:
            pos = win32gui.GetCursorPos()  # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
            win32gui.SetForegroundWindow(hwnd)
            win32gui.TrackPopupMenu(self.lmenu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, hwnd, None)
            win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
        elif lparam==win32con.WM_RBUTTONUP:
            pos = win32gui.GetCursorPos()  # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
            win32gui.SetForegroundWindow(hwnd)
            win32gui.TrackPopupMenu(self.rmenu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, hwnd, None)
            win32gui.PostMessage(hwnd, win32con.WM_NULL, 0, 0)
        elif lparam==win32con.WM_LBUTTONDBLCLK:
            nid = (hwnd, 0)
            win32gui.DestroyWindow(hwnd)
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
            win32gui.PostQuitMessage(0)
        return True  
    def OnCommand(self, hwnd, wparam, lparam):
        if lparam == self.cfrom:
            if win32con.CBN_SELCHANGE == win32gui.HIWORD(wparam): #LOWORD(wparam) == ID of combobox
                win32gui.SetWindowText(self.cname, win32gui.GetWindowText(self.cfrom))
        if wparam:
            if wparam == CID_REFR:
                self.createPopupMenu()
            elif wparam == CID_EXIT:
                win32gui.DestroyWindow(hwnd)
            elif wparam in self.cmdlist.keys():
                if self.cmdlist[wparam][0] == OPT_OPEN:
                    win32api.ShellExecute(win32con.NULL, "open", self.cmdlist[wparam][1], "", "", win32con.SW_SHOW);
                elif self.cmdlist[wparam][0] == OPT_COPY:
                    self.switchMode('c', self.cmdlist[wparam][1])
                    self.updated = False
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                elif self.cmdlist[wparam][0] == OPT_PAST:
                    self.switchMode('p', self.cmdlist[wparam][1])
                    self.updated = False
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                elif self.cmdlist[wparam][0] == OPT_SAVE:
                    win32clipboard.OpenClipboard()
                    fmt = win32clipboard.EnumClipboardFormats(0)
                    while fmt:
                        if fmt == win32con.CF_HDROP:
                            break
                        fmt = win32clipboard.EnumClipboardFormats(fmt)
                    else:
                        win32clipboard.CloseClipboard()
                        win32gui.MessageBox(hwnd, "未检测到可以粘贴的文件（夹）", "错误", win32con.MB_OK)
                        return
                    handles = win32clipboard.GetClipboardData(win32con.CF_HDROP)
                    win32clipboard.EmptyClipboard()
                    win32clipboard.CloseClipboard()
                    staticMethod.copyFiles(handles, self.cmdlist[wparam][1], True)
                    win32gui.MessageBox(win32con.NULL, "存档已完成！", "提示", win32con.MB_OK)
                    self.createPopupMenu()
        elif lparam:
            if lparam == self.psubf:
                win32gui.EnableWindow(self.psubs, True)
                win32gui.EnableWindow(self.pname, False)
            elif lparam == self.pnewf:
                win32gui.EnableWindow(self.psubs, False)
                win32gui.EnableWindow(self.pname, True)
            elif lparam == self.pcurf:
                win32gui.EnableWindow(self.psubs, False)
                win32gui.EnableWindow(self.pname, False)
            elif lparam == self.pself:
                cs = create_string_buffer(1024)
                if self.dll.chooseDir(self.wkdir.encode('GBK'), pointer(cs), "请选择将要粘贴到的目录：".encode('GBK')):
                    self.sldir = cs.value.decode("GBK")
                    win32gui.SetWindowText(self.pself, "指定文件夹（" + os.path.split(self.sldir)[1] + "）")
                else:
                    win32gui.SendMessage(self.pcurf, win32con.BM_SETCHECK, 1, 0)  #win32gui.SendMessage(self.csubs, win32con.CB_GETCOUNT, 1, 0)
                    win32gui.SendMessage(self.pself, win32con.BM_SETCHECK, 0, 0)
                    win32gui.SetWindowText(self.pself, "指定文件夹")
                win32gui.EnableWindow(self.psubs, False)
                win32gui.EnableWindow(self.pname, False)
            elif lparam == self.pfirm:
                win32clipboard.OpenClipboard()
                fmt = win32clipboard.EnumClipboardFormats(0)
                while fmt:
                    if fmt == win32con.CF_HDROP:
                        break
                    fmt = win32clipboard.EnumClipboardFormats(fmt)
                else:
                    win32clipboard.CloseClipboard()
                    win32gui.MessageBox(hwnd, "未检测到可以粘贴的文件（夹）", "错误", win32con.MB_OK)
                    return
                handles = win32clipboard.GetClipboardData(win32con.CF_HDROP)
                win32clipboard.EmptyClipboard()
                win32clipboard.CloseClipboard()
                if win32gui.SendMessage(self.pcurf, win32con.BM_GETCHECK, 0, 0):
                    staticMethod.copyFiles(handles, self.wkdir, win32gui.SendMessage(self.pmove, win32con.BM_GETCHECK, 0, 0))
                elif win32gui.SendMessage(self.psubf, win32con.BM_GETCHECK, 0, 0):
                    if len(win32gui.GetWindowText(self.psubs)):
                        staticMethod.copyFiles(handles, os.path.join(self.wkdir, win32gui.GetWindowText(self.psubs)), win32gui.SendMessage(self.pmove, win32con.BM_GETCHECK, 0, 0))
                    else:
                        return False
                elif win32gui.SendMessage(self.pnewf, win32con.BM_GETCHECK, 0, 0):
                    if len(win32gui.GetWindowText(self.pname)):
                        path = os.path.join(self.wkdir, win32gui.GetWindowText(self.pname))
                        if createDirAsPath(path):
                            staticMethod.copyFiles(handles, path, win32gui.SendMessage(self.pmove, win32con.BM_GETCHECK, 0, 0))
                        else:
                            return False
                    else:
                        return False
                elif win32gui.SendMessage(self.pself, win32con.BM_GETCHECK, 0, 0):
                    staticMethod.copyFiles(handles, self.sldir, win32gui.SendMessage(self.pmove, win32con.BM_GETCHECK, 0, 0))
                win32gui.MessageBox(win32con.NULL, "操作已完成！", "提示", win32con.MB_OK)
                self.updated = True
            elif lparam == self.cself:
                cs = create_string_buffer(1024)
                if self.dll.chooseDir(self.wkdir.encode('GBK'), pointer(cs), "请选择将要复制到的目录：".encode('GBK')):
                    self.sldir = cs.value.decode("GBK")
                    win32gui.SetWindowText(self.cself, "到指定的文件夹（" + os.path.split(self.sldir)[1] + ")")
                else:
                    win32gui.SendMessage(self.ccurf, win32con.BM_SETCHECK, 1, 0)  #win32gui.SendMessage(self.csubs, win32con.CB_GETCOUNT, 1, 0)
                    win32gui.SendMessage(self.cself, win32con.BM_SETCHECK, 0, 0)
                    win32gui.SetWindowText(self.cself, "到指定的文件夹")
            elif lparam == self.cfirm:
                oldname, newname = win32gui.GetWindowText(self.cfrom), win32gui.GetWindowText(self.cname)
                if len(oldname) and len(newname):
                    if getExtFromPath(oldname) != getExtFromPath(newname):
                        newname += getExtFromPath(oldname)
                    if win32gui.SendMessage(self.ccurf, win32con.BM_GETCHECK, 0, 0):
                        fpath = os.path.join(self.wkdir, oldname)
                        tpath = os.path.join(self.wkdir, newname)
                    elif win32gui.SendMessage(self.cself, win32con.BM_GETCHECK, 0, 0):
                        fpath = os.path.join(self.wkdir, oldname)
                        tpath = os.path.join(self.sldir, newname)
                    else:
                        return False
                    move = win32gui.SendMessage(self.cmove, win32con.BM_GETCHECK, 0, 0)
                    if staticMethod.copyFile(fpath, tpath, move):
                        if move:
                            lsts = createTreeAsPath(self.wkdir, self.regul, False, False, len(self.wkdir), True)
                            win32gui.SendMessage(self.cfrom, win32con.CB_RESETCONTENT, 0, 0)
                            for lst in lsts:
                                win32gui.SendMessage(self.cfrom, win32con.CB_ADDSTRING, 0, lst)
                    else:
                        return False
                    win32gui.MessageBox(win32con.NULL, "操作已完成！", "提示", win32con.MB_OK)
                    self.updated = True

class WindowApplication():
    def __init__(self, wndclassname, wndtitlename):
        self.func = WindowFunction(self)
        self.hwnd = self.func.prepareWindow(wndclassname, wndtitlename)
        self.__createApplication()
    def __createApplication(self):
        self.func.createTaskBar()
        self.func.createPopupMenu()
    def startHandleMessages(self):
        win32gui.PumpMessages()
    def showWindow(self, mode):
        win32gui.UpdateWindow(self.hwnd)
        win32gui.ShowWindow(self.hwnd, mode)

staticMethod.main()

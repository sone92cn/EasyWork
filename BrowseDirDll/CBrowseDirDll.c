#include <windows.h>
#include <ShlObj.h>

char* default_dir;

INT CALLBACK  BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData){
	switch(uMsg){
	case BFFM_INITIALIZED:
	   SendMessage(hwnd, BFFM_SETSELECTION, TRUE, (LPARAM)default_dir);
	   break;
	}
	return 0;
}

INT __cdecl chooseDir(char* szInitDir, char* szPathName, char* szInfo){  //__stdcall or __cdecl
	BROWSEINFO bInfo={0};
	LPITEMIDLIST lpDlist;
	default_dir = szInitDir;
	bInfo.hwndOwner = GetForegroundWindow();//父窗口
	bInfo.lpszTitle = szInfo; //"Select a folder to save:";
	bInfo.ulFlags = BIF_RETURNONLYFSDIRS |BIF_USENEWUI;
	bInfo.lpfn = BrowseCallbackProc;
	lpDlist = SHBrowseForFolder(&bInfo);
	if (NULL != lpDlist){
		SHGetPathFromIDList(lpDlist, szPathName); //MessageBox(NULL, szPathName, "NULL", MB_OK);
		return TRUE;
	}else{
		return FALSE;
	}
}

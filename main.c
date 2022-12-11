#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <uxtheme.h>
#include <locale.h>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <math.h>

#include <sqlext.h>
#include <sqltypes.h>
#include <sql.h>

#define LVS_EX_AUTOSIZECOLUMNS 0x10000000

#define WMU_UPDATE_GRID        WM_USER + 1
#define WMU_UPDATE_CACHE       WM_USER + 2
#define WMU_UPDATE_FILTER_SIZE WM_USER + 3
#define WMU_SET_HEADER_FILTERS WM_USER + 4
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 5
#define WMU_SET_CURRENT_CELL   WM_USER + 6
#define WMU_RESET_CACHE        WM_USER + 7
#define WMU_SET_FONT           WM_USER + 8
#define WMU_SET_THEME          WM_USER + 9
#define WMU_HIDE_COLUMN        WM_USER + 10
#define WMU_SHOW_COLUMNS       WM_USER + 11
#define WMU_HOT_KEYS           WM_USER + 12  
#define WMU_HOT_CHARS          WM_USER + 13

#define IDC_MAIN               100
#define IDC_TABLELIST          101
#define IDC_GRID               102
#define IDC_STATUSBAR          103
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROWS          5001
#define IDM_COPY_COLUMN        5002
#define IDM_FILTER_ROW         5003
#define IDM_HEADER_ROW         5004
#define IDM_DARK_THEME         5005
#define IDM_HIDE_COLUMN        5020

#define SB_VERSION             0
#define SB_TABLE_COUNT         1
#define SB_VIEW_COUNT          2
#define SB_TYPE                3
#define SB_ROW_COUNT           4
#define SB_CURRENT_CELL        5
#define SB_AUXILIARY           6

#define SPLITTER_WIDTH         5
#define MAX_TEXT_LENGTH        32000
#define MAX_DATA_LENGTH        32000
#define MAX_TABLE_LENGTH       2000
#define MAX_COLUMN_LENGTH      2000
#define MAX_ERROR_LENGTH       2000

#define ODBC_UNKNOWN           0
#define ODBC_ACCESS            1
#define ODBC_EXCEL             2
#define ODBC_EXCELX            3

#define APP_NAME               TEXT("odbc-wlx")
#define APP_VERSION            TEXT("1.0.2")

#define LCS_FINDFIRST          1
#define LCS_MATCHCASE          2
#define LCS_WHOLEWORDS         4
#define LCS_BACKWARDS          8

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

HWND getMainWindow(HWND hWnd);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
TCHAR* extractUrl(TCHAR* data);
void setClipboardText(const TCHAR* text);
BOOL isNumber(TCHAR* val);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	if (ul_reason_for_call == DLL_PROCESS_ATTACH && iniPath[0] == 0) {
		TCHAR path[MAX_PATH + 1] = {0};
		GetModuleFileName(hModule, path, MAX_PATH);
		TCHAR* dot = _tcsrchr(path, TEXT('.'));
		_tcsncpy(dot, TEXT(".ini"), 5);
		if (_taccess(path, 0) == 0)
			_tcscpy(iniPath, path);	
	}
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	snprintf(DetectString, maxlen, "MULTIMEDIA & (ext=\"MDB\" | ext=\"XLS\" | ext=\"XLSX\" | ext=\"XLSB\" | ext=\"XLSM\" | ext=\"DSN\")");
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

int __stdcall ListSearchTextW(HWND hWnd, TCHAR* searchString, int searchParameter) {
	HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
	HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);	
	
	TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
	int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
	int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));

	BOOL isFindFirst = searchParameter & LCS_FINDFIRST;		
	BOOL isBackward = searchParameter & LCS_BACKWARDS;
	BOOL isMatchCase = searchParameter & LCS_MATCHCASE;
	BOOL isWholeWords = searchParameter & LCS_WHOLEWORDS;	

	if (isFindFirst) {
		*(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) = 0;
		*(int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")) = 0;	
		*(int*)GetProp(hWnd, TEXT("CURRENTROWNO")) = isBackward ? rowCount - 1 : 0;
	}	
		
	int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
	int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
	int *pStartPos = (int*)GetProp(hWnd, TEXT("SEARCHCELLPOS"));	
	rowNo = rowNo == -1 || rowNo >= rowCount ? 0 : rowNo;
	colNo = colNo == -1 || colNo >= colCount ? 0 : colNo;	
			
	int pos = -1;
	do {
		for (; (pos == -1) && colNo < colCount; colNo++) {
			pos = findString(cache[rowNo][colNo] + *pStartPos, searchString, isMatchCase, isWholeWords);
			if (pos != -1) 
				pos += *pStartPos;			
			*pStartPos = pos == -1 ? 0 : pos + *pStartPos + _tcslen(searchString);
		}
		colNo = pos != -1 ? colNo - 1 : 0;
		rowNo += pos != -1 ? 0 : isBackward ? -1 : 1; 	
	} while ((pos == -1) && (isBackward ? rowNo > 0 : rowNo < rowCount));
	ListView_SetItemState(hGridWnd, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);

	TCHAR buf[256] = {0};
	if (pos != -1) {
		ListView_EnsureVisible(hGridWnd, rowNo, FALSE);
		ListView_SetItemState(hGridWnd, rowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
		
		TCHAR* val = cache[rowNo][colNo];
		int len = _tcslen(searchString);
		_sntprintf(buf, 255, TEXT("%ls%.*ls%ls"),
			pos > 0 ? TEXT("...") : TEXT(""), 
			len + pos + 10, val + pos,
			_tcslen(val + pos + len) > 10 ? TEXT("...") : TEXT(""));
		SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
	} else { 
		MessageBox(hWnd, searchString, TEXT("Not found:"), MB_OK);
	}
	SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)buf);	
	SetFocus(hGridWnd);	

	return 0;
}

int __stdcall ListSearchText(HWND hWnd, char* searchString, int searchParameter) {
	DWORD len = MultiByteToWideChar(CP_ACP, 0, searchString, -1, NULL, 0);
	TCHAR* searchString16 = (TCHAR*)calloc (len, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, searchString, -1, searchString16, len);
	int rc = ListSearchTextW(hWnd, searchString16, searchParameter);
	free(searchString16);
	return rc;
}

HWND APIENTRY ListLoadW (HWND hListerWnd, TCHAR* fileToLoad, int showFlags) {
	TCHAR* fileext = _tcsrchr(fileToLoad, TEXT('.'));
	_tcslwr(fileext);

	TCHAR* dbpath = calloc(MAX_PATH, sizeof(TCHAR));
	_tcsncpy(dbpath, fileToLoad, MAX_PATH);

	int odbcType = ODBC_UNKNOWN;
	TCHAR connectionString[MAX_TEXT_LENGTH] = {0};
	TCHAR connectionString2[MAX_TEXT_LENGTH] = {0};	 
	if (_tcscmp(fileext, TEXT(".mdb")) == 0 || _tcscmp(fileext, TEXT(".accdb")) == 0) {
		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Access Driver (*.mdb)};Dbq=%ls;Uid=Admin;Pwd=;ReadOnly=1;"), fileToLoad);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=%ls;Uid=Admin;Pwd=;ReadOnly=1;"), fileToLoad);
		odbcType = ODBC_ACCESS;
	} else if (_tcscmp(fileext, TEXT(".xls")) == 0) {
		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls)};Dbq=%ls;ReadOnly=1;"), fileToLoad);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=%ls;ReadOnly=1;"), fileToLoad);
		odbcType = ODBC_EXCEL;
	} else if (_tcscmp(fileext, TEXT(".xlsx")) == 0 || _tcscmp(fileext, TEXT(".xlsb")) == 0 || _tcscmp(fileext, TEXT(".xlsm")) == 0) {
		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=%ls;ReadOnly=1;"), fileToLoad);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=%ls;ReadOnly=1;"), fileToLoad);
		odbcType = ODBC_EXCELX;
	} else if (_tcscmp(fileext, TEXT(".dsn")) == 0) {
		TCHAR buf[32000];

		int len = GetPrivateProfileString(TEXT("ODBC"), NULL, NULL, buf, 32000, fileToLoad);
		int start = 0;
		for (int i = 0; i < len; i++) {
			if (buf[i] != 0)
				continue;

			TCHAR key[i - start + 1];
			_tcsncpy(key, buf + start, i - start + 1);
			TCHAR value[1024];
			GetPrivateProfileString(TEXT("ODBC"), key, NULL, value, 1024, fileToLoad);
			TCHAR pair[2000];
			BOOL isQ = _tcschr(value, TEXT(' ')) != 0;
			_sntprintf(pair, 2000, TEXT("%ls=%ls%ls%ls;"), key, isQ ? TEXT("{") : TEXT(""), value, isQ ? TEXT("}") : TEXT(""));
			_tcscat(connectionString, pair);

			start = i + 1;

			_tcslwr(key);
			_tcslwr(value);
			if (_tcscmp(key, TEXT("driver")) == 0)
				odbcType = _tcsstr(value, TEXT("*.mdb")) ? ODBC_ACCESS : 
					_tcsstr(value, TEXT("*.xls")) ? ODBC_EXCEL :
					_tcsstr(value, TEXT("*.xlsx")) ? ODBC_EXCELX :					
					ODBC_UNKNOWN;
			if (_tcscmp(key, TEXT("dbq")) == 0)
				_tcsncpy(dbpath, value, MAX_PATH);
		}
	}
	SQLHANDLE hEnv = 0;
	SQLHANDLE hConn = 0;
	SQLAllocHandle(SQL_HANDLE_ENV, SQL_NULL_HANDLE, &hEnv);
	SQLSetEnvAttr(hEnv, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0);
	SQLAllocHandle(SQL_HANDLE_DBC, hEnv, &hConn);

	if (!hEnv || !hConn)
		return 0;

	if (SQL_ERROR == SQLDriverConnect(hConn, NULL, connectionString, _tcslen(connectionString), 0, 0, NULL, SQL_DRIVER_NOPROMPT) &&
		SQL_ERROR == SQLDriverConnect(hConn, NULL, connectionString2, _tcslen(connectionString2), 0, 0, NULL, SQL_DRIVER_NOPROMPT)) {
		MessageBox(hListerWnd, TEXT("Can't connect to database"), NULL, MB_OK);
		SQLFreeHandle(SQL_HANDLE_DBC, hConn);
		SQLFreeHandle(SQL_HANDLE_ENV, hEnv);
		free(dbpath);
		return 0;
	}

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	
	setlocale(LC_CTYPE, "");

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindow(WC_STATIC, APP_NAME, WS_CHILD | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);	
	
	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("FILTERROW"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("HEADERROW"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DBPATH"), dbpath);	
	SetProp(hMainWnd, TEXT("TABLENAME"), calloc(MAX_TABLE_LENGTH, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("WHERE"), calloc(MAX_TEXT_LENGTH, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DBENV"), hEnv);
	SetProp(hMainWnd, TEXT("DB"), hConn);
	SetProp(hMainWnd, TEXT("CURRENTROWNO"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCOLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SEARCHCELLPOS"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ODBCTYPE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERPOSITION"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERALIGN"), calloc(1, sizeof(int)));	

	SetProp(hMainWnd, TEXT("DARKTHEME"), calloc(1, sizeof(int)));			
	SetProp(hMainWnd, TEXT("TEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR2"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERBACKCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("CURRENTCELLCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SELECTIONTEXTCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SELECTIONBACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERCOLOR"), calloc(1, sizeof(int)));		

	*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) = getStoredValue(TEXT("header-row"), 1);
	*(int*)GetProp(hMainWnd, TEXT("SPLITTERPOSITION")) = getStoredValue(TEXT("splitter-position"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 1);	
	*(int*)GetProp(hMainWnd, TEXT("ODBCTYPE")) = odbcType;
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);	

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);
	int sizes[7] = {35 * z, 110 * z, 180 * z, 225 * z, 420 * z, 500 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 7, (LPARAM)&sizes);
	
	HWND hListWnd = CreateWindow(TEXT("LISTBOX"), NULL, WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_TABSTOP | WS_HSCROLL | LBS_SORT,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_TABLELIST, GetModuleHandle(0), NULL);
	SetProp(hListWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hListWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
		
	int noLines = getStoredValue(TEXT("disable-grid-lines"), 0);	
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | (noLines ? 0 : LVS_EX_GRIDLINES) | LVS_EX_LABELTIP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));

	HMENU hGridMenu = CreatePopupMenu();
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_ROWS, TEXT("Copy row(s)"));
	AppendMenu(hGridMenu, MF_STRING, IDM_COPY_COLUMN, TEXT("Copy column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);
	AppendMenu(hGridMenu, MF_STRING, IDM_HIDE_COLUMN, TEXT("Hide column"));	
	AppendMenu(hGridMenu, MF_STRING, 0, NULL);	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_FILTER_ROW, TEXT("Filters"));		
	if (odbcType == ODBC_EXCEL || odbcType == ODBC_EXCELX) 
		AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("HEADERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_HEADER_ROW, TEXT("Header row"));	
	AppendMenu(hGridMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));			
	SetProp(hMainWnd, TEXT("GRIDMENU"), hGridMenu);

	int tCount = 0, vCount = 0;
	SQLHANDLE hStmt = 0;
	SQLAllocHandle(SQL_HANDLE_STMT, hConn, &hStmt);
	SQLTables(hStmt, NULL, 0, NULL, 0, NULL, 0, NULL, 0);

	SQLLEN res = 0;
	while (SQLFetch(hStmt) == SQL_SUCCESS) {
		SQLWCHAR tblName[MAX_DATA_LENGTH + 1];
		SQLGetData(hStmt, 3, SQL_WCHAR, tblName, MAX_DATA_LENGTH * sizeof(TCHAR), &res);

		SQLWCHAR tblType[MAX_DATA_LENGTH + 1];
		SQLGetData(hStmt, 4, SQL_WCHAR, tblType, MAX_DATA_LENGTH * sizeof(TCHAR), &res);
		_tcslwr(tblType);

		BOOL isSystem = _tcscmp(tblType, TEXT("system table")) == 0;		
		if (odbcType == ODBC_ACCESS && isSystem)
			continue;

		if (odbcType == ODBC_EXCEL || odbcType == ODBC_EXCELX) {
			TCHAR* tail = _tcsstr(tblName, TEXT("$'"));	
			int len = tail ? _tcslen(tail) : 0;
			if (len > 2 || len < 2 && !isSystem)
				continue;
			
			// Remove $-tail or quotes e.g. 'tblName$'
			TCHAR tmpName[MAX_DATA_LENGTH] = {0};
			for (int i = 0; i < _tcslen(tblName) && tblName[i] != TEXT('$'); i++)
				tmpName[i] = tblName[i];	
			_sntprintf(tblName, MAX_DATA_LENGTH, TEXT("%ls"), tmpName + (tmpName[0] == TEXT('\'')));	
		}

		int pos = ListBox_AddString(hListWnd, tblName);
		BOOL isTable = _tcsstr(tblType, TEXT("view")) == 0;
		SendMessage(hListWnd, LB_SETITEMDATA, pos, isTable);

		tCount += isTable;
		vCount += !isTable;
	}
	SQLCloseCursor(hStmt);
	SQLFreeHandle(SQL_HANDLE_STMT, hStmt);

	TCHAR buf[255];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);	
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);
	_sntprintf(buf, 255, TEXT(" Tables: %i"), tCount);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_TABLE_COUNT, (LPARAM)buf);
	_sntprintf(buf, 255, TEXT(" Views: %i"), vCount);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VIEW_COUNT, (LPARAM)buf);
	
	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WMU_SET_THEME, 0, 0);	
	ListBox_SetCurSel(hListWnd, 0);
	SendMessage(hMainWnd, WMU_UPDATE_GRID, 0, 0);
	ShowWindow(hMainWnd, SW_SHOW);
	SetFocus(hListWnd);	

	return hMainWnd;
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* fileToLoadW = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, fileToLoadW, size);
	HWND hWnd = ListLoadW(hListerWnd, fileToLoadW, showFlags);
	free(fileToLoadW);
	return hWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-position"), *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("filter-row"), *(int*)GetProp(hWnd, TEXT("FILTERROW")));		
	setStoredValue(TEXT("header-row"), *(int*)GetProp(hWnd, TEXT("HEADERROW")));	
	setStoredValue(TEXT("dark-theme"), *(int*)GetProp(hWnd, TEXT("DARKTHEME")));
	
	SQLHANDLE hEnv = (SQLHANDLE)GetProp(hWnd, TEXT("DBENV"));
	SQLHANDLE hConn = (SQLHANDLE)GetProp(hWnd, TEXT("DB"));
	SQLDisconnect(hConn);
	SQLFreeHandle(SQL_HANDLE_DBC, hConn);
	SQLFreeHandle(SQL_HANDLE_ENV, hEnv);

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	free((int*)GetProp(hWnd, TEXT("FILTERROW")));
	free((int*)GetProp(hWnd, TEXT("HEADERROW")));
	free((int*)GetProp(hWnd, TEXT("DARKTHEME")));
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((TCHAR*)GetProp(hWnd, TEXT("DBPATH")));	
	free((TCHAR*)GetProp(hWnd, TEXT("TABLENAME")));
	free((TCHAR*)GetProp(hWnd, TEXT("WHERE")));
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	free((int*)GetProp(hWnd, TEXT("FONTSIZE")));
	free((int*)GetProp(hWnd, TEXT("FILTERCOLOR")));		
	free((int*)GetProp(hWnd, TEXT("CURRENTROWNO")));				
	free((int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));
	free((int*)GetProp(hWnd, TEXT("SEARCHCELLPOS")));
	free((int*)GetProp(hWnd, TEXT("ODBCTYPE")));
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	free((int*)GetProp(hWnd, TEXT("FILTERALIGN")));
		
	free((int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR2")));	
	free((int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")));
	free((int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")));	

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));	
	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("GRIDMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("FILTERROW"));	
	RemoveProp(hWnd, TEXT("HEADERROW"));	
	RemoveProp(hWnd, TEXT("DARKTHEME"));	
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("DBENV"));
	RemoveProp(hWnd, TEXT("DB"));
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("DBPATH"));	
	RemoveProp(hWnd, TEXT("TABLENAME"));
	RemoveProp(hWnd, TEXT("WHERE"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("SPLITTERPOSITION"));
	RemoveProp(hWnd, TEXT("CURRENTROWNO"));	
	RemoveProp(hWnd, TEXT("CURRENTCOLNO"));
	RemoveProp(hWnd, TEXT("SEARCHCELLPOS"));
	RemoveProp(hWnd, TEXT("ODBCTYPE"));
	RemoveProp(hWnd, TEXT("FILTERALIGN"));
	RemoveProp(hWnd, TEXT("LASTFOCUS"));
	
	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));
	RemoveProp(hWnd, TEXT("TEXTCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR2"));	
	RemoveProp(hWnd, TEXT("FILTERTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERBACKCOLOR"));
	RemoveProp(hWnd, TEXT("CURRENTCELLCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("SELECTIONBACKCOLOR"));
	RemoveProp(hWnd, TEXT("SPLITTERCOLOR"));
	RemoveProp(hWnd, TEXT("BACKBRUSH"));
	RemoveProp(hWnd, TEXT("FILTERBACKBRUSH"));		
	RemoveProp(hWnd, TEXT("SPLITTERBRUSH"));
	RemoveProp(hWnd, TEXT("GRIDMENU"));

	DestroyWindow(hWnd);
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int splitterW = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			GetClientRect(hWnd, &rc);
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SetWindowPos(hListWnd, 0, 0, 0, splitterW, rc.bottom - statusH, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hGridWnd, 0, splitterW + SPLITTER_WIDTH, 0, rc.right - splitterW - SPLITTER_WIDTH, rc.bottom - statusH, SWP_NOZORDER);
		}
		break;

		case WM_PAINT: {
			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			rc.left = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			rc.right = rc.left + SPLITTER_WIDTH;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("SPLITTERBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;
		
		case WM_SETCURSOR: {
			SetCursor(LoadCursor(0, GetProp(hWnd, TEXT("ISMOUSEHOVER")) ? IDC_SIZEWE : IDC_ARROW));
			return TRUE;
		}
		break;
		
		case WM_SETFOCUS: {
			SetFocus(GetProp(hWnd, TEXT("LASTFOCUS")));
		}
		break;				

		case WM_LBUTTONDOWN: {
			int x = GET_X_LPARAM(lParam);
			int pos = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			if (x >= pos && x <= pos + SPLITTER_WIDTH) {
				SetProp(hWnd, TEXT("ISMOUSEDOWN"), (HANDLE)1);
				SetCapture(hWnd);
			}
			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
		}
		break;
		
		case WM_MOUSEMOVE: {
			DWORD x = GET_X_LPARAM(lParam);
			int* pPos = (int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			
			if (!GetProp(hWnd, TEXT("ISMOUSEHOVER")) && *pPos <= x && x <= *pPos + SPLITTER_WIDTH) {
				TRACKMOUSEEVENT tme = {sizeof(TRACKMOUSEEVENT), TME_LEAVE, hWnd, 0};
				TrackMouseEvent(&tme);	
				SetProp(hWnd, TEXT("ISMOUSEHOVER"), (HANDLE)1);
			}
			
			if (GetProp(hWnd, TEXT("ISMOUSEDOWN")) && x > 0 && x < 32000) {
				*pPos = x;
				SendMessage(hWnd, WM_SIZE, 0, 0);
			}
		}
		break;
		
		case WM_MOUSELEAVE: {
			SetProp(hWnd, TEXT("ISMOUSEHOVER"), 0);
		}
		break;
		
		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
			}
		}
		break;
		
		case WM_KEYDOWN: {
			if (SendMessage(hWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
		
		case WM_CTLCOLORLISTBOX: {
			SetBkColor((HDC)wParam, *(int*)GetProp(hWnd, TEXT("BACKCOLOR")));
			SetTextColor((HDC)wParam, *(int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
			return (INT_PTR)(HBRUSH)GetProp(hWnd, TEXT("BACKBRUSH"));	
		}
		break;
				
		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDC_TABLELIST && HIWORD(wParam) == LBN_SELCHANGE)
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
				
			if (cmd == IDC_TABLELIST && HIWORD(wParam) == LBN_SETFOCUS) 
				SetProp(hWnd, TEXT("LASTFOCUS"), (HWND)lParam);	
			
			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROWS || cmd == IDM_COPY_COLUMN) {
				HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				
				int colCount = Header_GetItemCount(hHeader);
				int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
				int selCount = ListView_GetSelectedCount(hGridWnd);

				if (rowNo == -1 ||
					rowNo >= rowCount ||
					colCount == 0 ||
					cmd == IDM_COPY_CELL && colNo == -1 || 
					cmd == IDM_COPY_CELL && colNo >= colCount || 
					cmd == IDM_COPY_COLUMN && colNo == -1 || 
					cmd == IDM_COPY_COLUMN && colNo >= colCount || 					
					cmd == IDM_COPY_ROWS && selCount == 0) {
					setClipboardText(TEXT(""));
					return 0;
				}
						
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				TCHAR* delimiter = getStoredString(TEXT("column-delimiter"), TEXT("\t"));

				int len = 0;
				if (cmd == IDM_COPY_CELL) 
					len = _tcslen(cache[rowNo][colNo]);
								
				if (cmd == IDM_COPY_ROWS) {
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) 
								len += _tcslen(cache[rowNo][colNo]) + 1; /* column delimiter */
						}
													
						len++; /* \n */		
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
				}				

				if (cmd == IDM_COPY_COLUMN) {
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						len += _tcslen(cache[rowNo][colNo]) + 1 /* \n */;
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
					} 
				}	

				TCHAR* buf = calloc(len + 1, sizeof(TCHAR));
				if (cmd == IDM_COPY_CELL)
					_tcscat(buf, cache[rowNo][colNo]);
								
				if (cmd == IDM_COPY_ROWS) {
					int pos = 0;
					int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1) {
						for (int colNo = 0; colNo < colCount; colNo++) {
							if (ListView_GetColumnWidth(hGridWnd, colNo)) {
								int len = _tcslen(cache[rowNo][colNo]);
								_tcsncpy(buf + pos, cache[rowNo][colNo], len);
								buf[pos + len] = delimiter[0];
								pos += len + 1;
							}
						}

						buf[pos - (pos > 0)] = TEXT('\n');
						rowNo = ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);	
					}
					buf[pos - 1] = 0; // remove last \n
				}				

				if (cmd == IDM_COPY_COLUMN) {
					int pos = 0;
					int rowNo = selCount < 2 ? 0 : ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
					while (rowNo != -1 && rowNo < rowCount) {
						int len = _tcslen(cache[rowNo][colNo]);
						_tcsncpy(buf + pos, cache[rowNo][colNo], len);
						rowNo = selCount < 2 ? rowNo + 1 : ListView_GetNextItem(hGridWnd, rowNo, LVNI_SELECTED);
						if (rowNo != -1 && rowNo < rowCount)
							buf[pos + len] = TEXT('\n');
						pos += len + 1;								
					} 
				}
									
				setClipboardText(buf);
				free(buf);
				free(delimiter);
			}
			
			if (cmd == IDM_HIDE_COLUMN) {
				int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
				SendMessage(hWnd, WMU_HIDE_COLUMN, colNo, 0);
			}						
			
			if (cmd == IDM_FILTER_ROW || cmd == IDM_HEADER_ROW || cmd == IDM_DARK_THEME) {
				HMENU hMenu = (HMENU)GetProp(hWnd, TEXT("GRIDMENU"));
				int* pOpt = (int*)GetProp(hWnd, cmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : cmd == IDM_HEADER_ROW ? TEXT("HEADERROW") : TEXT("DARKTHEME"));
				*pOpt = (*pOpt + 1) % 2;
				Menu_SetItemState(hMenu, cmd, *pOpt ? MFS_CHECKED : 0);
				
				UINT msg = cmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : cmd == IDM_HEADER_ROW ? WMU_UPDATE_GRID : WMU_SET_THEME;
				SendMessage(hWnd, msg, 0, 0);				
			}				
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) {
				LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
				LV_ITEM* pItem= &(pDispInfo)->item;
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

				if(pItem->mask & LVIF_TEXT)
					pItem->pszText = cache[pItem->iItem][pItem->iSubItem];
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				// Hide or sort the column
				if (HIWORD(GetKeyState(VK_CONTROL))) {
					HWND hGridWnd = pHdr->hwndFrom;
					HWND hHeader = ListView_GetHeader(hGridWnd);
					int colNo = lv->iSubItem;
					
					HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
					SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, colNo));				
					ListView_SetColumnWidth(pHdr->hwndFrom, colNo, 0); 
					InvalidateRect(hHeader, NULL, TRUE);
				} else {
					int colNo = lv->iSubItem + 1;
					int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
					int orderBy = *pOrderBy;
					*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
					SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);				
				}				
			}

			if (pHdr->idFrom == IDC_GRID && (pHdr->code == (DWORD)NM_CLICK || pHdr->code == (DWORD)NM_RCLICK)) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				SendMessage(hWnd, WMU_SET_CURRENT_CELL, ia->iItem, ia->iSubItem);
			}
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_CLICK && HIWORD(GetKeyState(VK_MENU))) {	
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				
				TCHAR* url = extractUrl(cache[ia->iItem][ia->iSubItem]);
				ShellExecute(0, TEXT("open"), url, 0, 0 , SW_SHOW);
				free(url);
			}			

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
				POINT p;
				GetCursorPos(&p);
				TrackPopupMenu(GetProp(hWnd, TEXT("GRIDMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				NMLISTVIEW* lv = (NMLISTVIEW*)lParam;
				if (lv->uOldState != lv->uNewState && (lv->uNewState & LVIS_SELECTED))				
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, lv->iItem, *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")));	
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43) { // C
					BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
					BOOL isShift = HIWORD(GetKeyState(VK_SHIFT)); 
					BOOL isCopyColumn = getStoredValue(TEXT("copy-column"), 0) && ListView_GetSelectedCount(pHdr->hwndFrom) > 1;
					if (!isCtrl && !isShift)
						return FALSE;
						
					int action = !isShift && !isCopyColumn ? IDM_COPY_CELL : isCtrl || isCopyColumn ? IDM_COPY_COLUMN : IDM_COPY_ROWS;
					SendMessage(hWnd, WM_COMMAND, action, 0);

					SendMessage(hWnd, WM_COMMAND, action, 0);
					return TRUE;
				}

				if (kd->wVKey == 0x41 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + A
					HWND hGridWnd = pHdr->hwndFrom;
					SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));					
					ListView_SetItemState(hGridWnd, -1, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED);
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, rowNo, colNo);
					SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
					InvalidateRect(hGridWnd, NULL, TRUE);
				}
				
				if (kd->wVKey == 0x20 && HIWORD(GetKeyState(VK_CONTROL))) { // Ctrl + Space				
					SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);					
					return TRUE;
				}				
				
				if (kd->wVKey == VK_LEFT || kd->wVKey == VK_RIGHT) {
					int colCount = Header_GetItemCount(ListView_GetHeader(pHdr->hwndFrom));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO")) + (kd->wVKey == VK_RIGHT ? 1 : -1);
					colNo = colNo < 0 ? colCount - 1 : colNo > colCount - 1 ? 0 : colNo;
					SendMessage(hWnd, WMU_SET_CURRENT_CELL, *(int*)GetProp(hWnd, TEXT("CURRENTROWNO")), colNo);
					return TRUE;
				}
			}

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_GRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
				
			if (pHdr->code == (UINT)NM_SETFOCUS)
				SetProp(hWnd, TEXT("LASTFOCUS"), pHdr->hwndFrom);
			
			if (pHdr->idFrom == IDC_GRID && pHdr->code == (UINT)NM_CUSTOMDRAW) {
				int result = CDRF_DODEFAULT;
				
				NMLVCUSTOMDRAW* pCustomDraw = (LPNMLVCUSTOMDRAW)lParam;				
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_PREPAINT) 
					result = CDRF_NOTIFYITEMDRAW;
	
				if (pCustomDraw->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					if (ListView_GetItemState(pHdr->hwndFrom, pCustomDraw->nmcd.dwItemSpec, LVIS_SELECTED)) {
						pCustomDraw->nmcd.uItemState &= ~CDIS_SELECTED;
						result = CDRF_NOTIFYSUBITEMDRAW;
					} else {
						pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, pCustomDraw->nmcd.dwItemSpec % 2 == 0 ? TEXT("BACKCOLOR") : TEXT("BACKCOLOR2"));
					}				
				}
				
				if (pCustomDraw->nmcd.dwDrawStage == (CDDS_ITEMPREPAINT | CDDS_SUBITEM)) {
					int rowNo = *(int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
					int colNo = *(int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
					BOOL isCurrCell = (pCustomDraw->nmcd.dwItemSpec == (DWORD)rowNo) && (pCustomDraw->iSubItem == colNo);
					pCustomDraw->clrText = *(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR"));
					pCustomDraw->clrTextBk = *(int*)GetProp(hWnd, isCurrCell ? TEXT("CURRENTCELLCOLOR") : TEXT("SELECTIONBACKCOLOR"));
				}
	
				return result;
			}				
		}
		break;
		
		// wParam = colNo
		case WMU_HIDE_COLUMN: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);		
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colNo = (int)wParam;

			HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
			SetWindowLongPtr(hEdit, GWLP_USERDATA, (LONG_PTR)ListView_GetColumnWidth(hGridWnd, colNo));				
			ListView_SetColumnWidth(hGridWnd, colNo, 0); 
			InvalidateRect(hHeader, NULL, TRUE);			
		}
		break;	
		
		case WMU_SHOW_COLUMNS: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			for (int colNo = 0; colNo < colCount; colNo++) {
				if (ListView_GetColumnWidth(hGridWnd, colNo) == 0) {
					HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
					ListView_SetColumnWidth(hGridWnd, colNo, (int)GetWindowLongPtr(hEdit, GWLP_USERDATA));
				}
			}

			InvalidateRect(hGridWnd, NULL, TRUE);		
		}
		break;			

		case WMU_UPDATE_GRID: {
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SQLHANDLE hConn = (SQLHANDLE)GetProp(hWnd, TEXT("DB"));
			int odbcType = *(int*)GetProp(hWnd, TEXT("ODBCTYPE"));
			BOOL isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
			int filterAlign = *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));
			
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			ListView_SetItemCount(hGridWnd, 0);
			SendMessage(hWnd, WMU_SET_CURRENT_CELL, 0, 0);

			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++) 
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);

			TCHAR* tablename = (TCHAR*)GetProp(hWnd, TEXT("TABLENAME"));			
			int pos = ListBox_GetCurSel(hListWnd);
			if (odbcType == ODBC_EXCEL || odbcType == ODBC_EXCELX) {
				TCHAR tmpName[MAX_TABLE_LENGTH] = {0};
				ListBox_GetText(hListWnd, pos, tmpName);
				BOOL q = FALSE;
				for (int i = 0; i < _tcslen(tmpName) && !q; i++)
					q = !_istalnum(tmpName[i]) && tmpName[i] != TEXT('_');
				_sntprintf(tablename, MAX_TABLE_LENGTH, TEXT("%ls%ls$%ls"), q ? TEXT("'"): TEXT(""), tmpName, q ? TEXT("'"): TEXT(""));
			} else {
				ListBox_GetText(hListWnd, pos, tablename);
			}
			
			TCHAR buf[255];
			int type = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);
			_sntprintf(buf, 255, type ? TEXT(" TABLE"): TEXT("  VIEW"));
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TYPE, (LPARAM)buf);

			SQLHANDLE hStmt = 0;
			SQLAllocHandle(SQL_HANDLE_STMT, hConn, &hStmt);
			int len = 1024 + MAX_PATH + _tcslen(tablename);
			TCHAR query[len + 1];
			if (odbcType == ODBC_EXCEL || odbcType == ODBC_EXCELX) {
				TCHAR* dbpath = (TCHAR*)GetProp(hWnd, TEXT("DBPATH"));
				_sntprintf(query, len, TEXT("select * from \"Excel 8.0;HDR=%ls;IMEX=1;Database=%ls;\".\"%ls\" where 1 = 2"), 
					isHeaderRow ? TEXT("YES") : TEXT("NO"), dbpath, tablename);
			} else { 
				_sntprintf(query, len, TEXT("select * from \"%ls\" where 1 = 2"), tablename);
			}
			
			if (SQL_SUCCESS == SQLExecDirect(hStmt, query, SQL_NTS)) {
				SQLSMALLINT colCount = 0;
				SQLNumResultCols(hStmt, &colCount);

				for (int colNo = 0; colNo < colCount; colNo++) {
					SQLWCHAR colName[MAX_COLUMN_LENGTH];
					SQLSMALLINT colType = 0;
					SQLDescribeCol(hStmt, colNo + 1, colName, MAX_COLUMN_LENGTH, 0, &colType, 0, 0, 0);
					if (!isHeaderRow && (odbcType == ODBC_EXCEL || odbcType == ODBC_EXCELX))
						_sntprintf(colName, 64, TEXT("Column #%i"), colNo + 1);

					int fmt = colType == SQL_DECIMAL || colType == SQL_NUMERIC || colType == SQL_REAL || colType == SQL_FLOAT || colType == SQL_DOUBLE ||
						colType == SQL_SMALLINT || colType == SQL_INTEGER || colType == SQL_BIT || colType == SQL_TINYINT || colType == SQL_BIGINT ?
						LVCFMT_RIGHT :
						LVCFMT_LEFT;
						
					ListView_AddColumn(hGridWnd, colName, fmt);	
				}

				int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
				for (int colNo = 0; colNo < colCount; colNo++) {
					RECT rc;
					Header_GetItemRect(hHeader, colNo, &rc);
					HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_BORDER | WS_TABSTOP, 0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
					SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
					SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
				}
			} else {
				SQLWCHAR err[MAX_ERROR_LENGTH + 1];
				SQLWCHAR code[6];
				SQLGetDiagRec(SQL_HANDLE_STMT, hStmt, 1, code, NULL, err, MAX_ERROR_LENGTH, NULL);
				TCHAR msg[MAX_ERROR_LENGTH + 100];
				_sntprintf(msg, MAX_ERROR_LENGTH + 100, TEXT("Error (%ls): %ls"), code, err);
				MessageBox(hWnd, msg, NULL, MB_OK);
			}
			SQLCloseCursor(hStmt);
			SQLFreeHandle(SQL_HANDLE_STMT, hStmt);

			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;
			SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);

			SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_CACHE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SQLHANDLE hConn = (SQLHANDLE)GetProp(hWnd, TEXT("DB"));
			int odbcType = *(int*)GetProp(hWnd, TEXT("ODBCTYPE"));
			TCHAR* tablename = (TCHAR*)GetProp(hWnd, TEXT("TABLENAME"));
			TCHAR* where = (TCHAR*)GetProp(hWnd, TEXT("WHERE"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int orderBy = *(int*)GetProp(hWnd, TEXT("ORDERBY"));
			BOOL isHeaderRow = *(int*)GetProp(hWnd, TEXT("HEADERROW"));
			BOOL isExcel = odbcType == ODBC_EXCEL || odbcType == ODBC_EXCELX;
			
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			ListView_SetItemCount(hGridWnd, 0);

			_sntprintf(where, MAX_TEXT_LENGTH, TEXT("where (1 = 1)"));
			for (int colNo = 0; colNo < colCount; colNo++) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				int len = GetWindowTextLength(hEdit);
				if (len > 0) {
					TCHAR colName[256] = {0};
					if (!isHeaderRow && isExcel)
						_sntprintf(colName, 255, TEXT("F%i"), colNo + 1);
					else	
						Header_GetItemText(hHeader, colNo, colName, 255);

					TCHAR val[len + 1];
					GetWindowText(hEdit, val, len + 1);
					BOOL hasPrefix = len > 1 && (val[0] == TEXT('=') || val[0] == TEXT('!') || val[0] == TEXT('<') || val[0] == TEXT('>'));

					TCHAR qval[2 * len + 1];
					qval[0] = TEXT('\'');
					int pos = 1;
					for (int i = hasPrefix; i < len; i++) {
						qval[pos] = val[i];
						if (val[i] == TEXT('\'')) {
							pos++;
							qval[pos] = val[i];
						}
						pos++;
					}
					qval[pos] = TEXT('\'');
					qval[pos + 1] = 0;

					_tcscat(where, TEXT(" and \""));
					_tcscat(where, colName);

					TCHAR cond[MAX_TEXT_LENGTH];
					_sntprintf(cond, MAX_TEXT_LENGTH, len == 1 ? TEXT("\" like '%%' & %ls & '%%'") :
						val[0] == TEXT('=') ? TEXT("\" = %ls") :
						val[0] == TEXT('!') ? TEXT("\" not like '%%' & %ls & '%%'") :
						val[0] == TEXT('>') ? TEXT("\" > %ls") :
						val[0] == TEXT('<') ? TEXT("\" < %ls") :
						TEXT("\" like '%%' & %ls & '%%'"), isNumber(val + hasPrefix) ? val + hasPrefix : qval);

					_tcscat(where, cond);
				}
			}

			int cacheSize = 1000;
			TCHAR*** cache = calloc(cacheSize, sizeof(TCHAR*));
			int rowLimit = getStoredValue(TEXT("max-row-count"), 0);
			
			SQLHANDLE hStmt = 0;
			SQLAllocHandle(SQL_HANDLE_STMT, hConn, &hStmt);
			int len = 1024 + MAX_PATH + _tcslen(tablename) + _tcslen(where);
			TCHAR* query = calloc(len + 1, sizeof(TCHAR));
			
			TCHAR orderBy16[32] = {0};
			if (orderBy > 0)
				_sntprintf(orderBy16, 32, TEXT("order by %i"), orderBy);
			if (orderBy < 0)
				_sntprintf(orderBy16, 32, TEXT("order by %i desc"), -orderBy);
			
			if (isExcel) {
				TCHAR* dbpath = (TCHAR*)GetProp(hWnd, TEXT("DBPATH"));
				_sntprintf(query, len, TEXT("select * from \"Excel 8.0;HDR=%ls;IMEX=1;Database=%ls;\".\"%ls\" %ls %ls"), 
					isHeaderRow ? TEXT("YES") : TEXT("NO"), dbpath, tablename, where, orderBy16);
			} else { 
				_sntprintf(query, len, TEXT("select * from \"%ls\" %ls %ls"), tablename, where, orderBy16);
			}
						
			int rowNo = -1;
			if(SQL_SUCCESS == SQLExecDirect(hStmt, query, SQL_NTS)) {
				rowNo = 0;
				while (SQLFetch(hStmt) == SQL_SUCCESS && (rowNo < rowLimit || rowLimit == 0)) {
					if (rowNo >= cacheSize) {
						cacheSize += 100;
						cache = realloc(cache, cacheSize * sizeof(TCHAR**));
					}					
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));

					for (int colNo = 0; colNo < colCount; colNo++) {
						SQLLEN bytes = 0;
						SQLWCHAR val[MAX_DATA_LENGTH];
						SQLGetData(hStmt, colNo + 1, SQL_C_TCHAR, val, MAX_DATA_LENGTH * sizeof(TCHAR), &bytes);
						
						int len = bytes == -1 /* NULL */ ? 0 : bytes / 2;
						cache[rowNo][colNo] = calloc(len + 1, sizeof(TCHAR));
						
						if (len > 0) {
							// Excel: fix trailing zero .0
							if (isExcel && len > 2 && (val[len - 2] == TEXT('.')) && (val[len - 1] == TEXT('0'))) {
								BOOL isNum = TRUE;
								for (int i = 0; isNum && i < len - 2; i++)
									isNum = _istdigit(val[i]);
								len -= isNum ? 2 : 0;
							}
							
							_tcsncpy(cache[rowNo][colNo], val, len);
						}
					}
					
					rowNo++;
				}
				
				if (rowNo > 0)
					cache = realloc(cache, rowNo * sizeof(TCHAR**));
			}
			SQLCloseCursor(hStmt);
			SQLFreeHandle(SQL_HANDLE_STMT, hStmt);
			free(query);			
			
			if (rowNo > 0) {
				SetProp(hWnd, TEXT("CACHE"), cache);
			} else {
				free(cache);
			}
			
			if (_tcscmp(where, TEXT("where (1 = 1)")) == 0)
				*pTotalRowCount = rowNo != -1 ? rowNo : 0;
			*pRowCount = rowNo != -1 ? rowNo : 0;
			ListView_SetItemCount(hGridWnd, *pRowCount);
						
			TCHAR buf[1024];
			if (rowNo != -1)	
				_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), *pRowCount, *pTotalRowCount);
			else 
				_sntprintf(buf, 255, TEXT(" Rows: N/A"));
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, 0);			
		}
		break;

		case WMU_UPDATE_FILTER_SIZE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);
			for (int colNo = 0; colNo < colCount; colNo++) {
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left, h2, rc.right - rc.left, h2 + 1, SWP_NOZORDER);							
			}
		}
		break;
		
		case WMU_SET_HEADER_FILTERS: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
			int colCount = Header_GetItemCount(hHeader);
			
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
			LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
			styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
			SetWindowLongPtr(hHeader, GWL_STYLE, styles);

			for (int colNo = 0; colNo < colCount; colNo++) 		
				ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);

			// Bug fix: force Windows to redraw header
			SetWindowPos(hGridWnd, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOMOVE);
			SendMessage(getMainWindow(hWnd), WM_SIZE, 0, 0);
			
			if (isFilterRow)				
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);

			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;

		case WMU_AUTO_COLUMN_SIZE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hGridWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
				ListView_SetColumnWidth(hGridWnd, 0, 100);
				
			int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
				}
			}

			// Fix last column				
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hGridWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[MAX_COLUMN_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, name16, MAX_COLUMN_LENGTH);
				
				SIZE s = {0};
				HDC hDC = GetDC(hHeader);
				HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
				GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
				SelectObject(hDC, hOldFont);
				ReleaseDC(hHeader, hDC);

				int w = s.cx + 12;
				if (ListView_GetColumnWidth(hGridWnd, colNo) < w)
					ListView_SetColumnWidth(hGridWnd, colNo, w);
					
				if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);	
			}

			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hGridWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;
		
		// wParam = rowNo, lParam = colNo
		case WMU_SET_CURRENT_CELL: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);
						
			int *pRowNo = (int*)GetProp(hWnd, TEXT("CURRENTROWNO"));
			int *pColNo = (int*)GetProp(hWnd, TEXT("CURRENTCOLNO"));
			if (*pRowNo == wParam && *pColNo == lParam)
				return 0;
			
			RECT rc, rc2;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);			
			InvalidateRect(hGridWnd, &rc, TRUE);
			
			*pRowNo = wParam;
			*pColNo = lParam;
			ListView_GetSubItemRect(hGridWnd, *pRowNo, *pColNo, LVIR_BOUNDS, &rc);
			if (*pColNo == 0)
				rc.right = ListView_GetColumnWidth(hGridWnd, *pColNo);
			InvalidateRect(hGridWnd, &rc, FALSE);
			
			GetClientRect(hGridWnd, &rc2);
			int w = rc.right - rc.left;
			int dx = rc2.right < rc.right ? rc.left - rc2.right + w : rc.left < 0 ? rc.left : 0;

			ListView_Scroll(hGridWnd, dx, 0);
			
			TCHAR buf[32] = {0};
			if (*pColNo != - 1 && *pRowNo != -1)
				_sntprintf(buf, 32, TEXT(" %i:%i"), *pRowNo + 1, *pColNo + 1);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_CELL, (LPARAM)buf);
		}
		break;

		case WMU_RESET_CACHE: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));

			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			if (colCount > 0 && cache != 0) {
				for (int rowNo = 0; rowNo < *pRowCount; rowNo++) {
					if (cache[rowNo]) {
						for (int colNo = 0; colNo < colCount; colNo++)
							if (cache[rowNo][colNo])
								free(cache[rowNo][colNo]);

						free(cache[rowNo]);
					}
					cache[rowNo] = 0;
				}
				free(cache);
			}

			SetProp(hWnd, TEXT("CACHE"), 0);
			*pRowCount = 0;
		}
		break;
		
		// wParam - size delta
		case WMU_SET_FONT: {
			int* pFontSize = (int*)GetProp(hWnd, TEXT("FONTSIZE"));
			if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
				return 0;
			*pFontSize += wParam;
			DeleteFont(GetProp(hWnd, TEXT("FONT")));

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			SendMessage(hListWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);

			HWND hHeader = ListView_GetHeader(hGridWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

			int w = 0;
			HDC hDC = GetDC(hListWnd);
			HFONT hOldFont = (HFONT)SelectObject(hDC, hFont);
			TCHAR buf[MAX_TABLE_LENGTH]; 
			for (int i = 0; i < ListBox_GetCount(hListWnd); i++) {
				SIZE s = {0};
				ListBox_GetText(hListWnd, i, buf);
				GetTextExtentPoint32(hDC, buf, _tcslen(buf), &s);			
				if (w < s.cx)
					w = s.cx;
			}
			SelectObject(hDC, hOldFont);
			ReleaseDC(hHeader, hDC);
			SendMessage(hListWnd, LB_SETHORIZONTALEXTENT, w, 0);

			SetProp(hWnd, TEXT("FONT"), hFont);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;	
		
		case WMU_SET_THEME: {
			HWND hGridWnd = GetDlgItem(hWnd, IDC_GRID);
			BOOL isDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));
			
			int textColor = !isDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
			int backColor = !isDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
			int backColor2 = !isDark ? getStoredValue(TEXT("back-color2"), RGB(240, 240, 240)) : getStoredValue(TEXT("back-color2-dark"), RGB(52, 52, 52));
			int filterTextColor = !isDark ? getStoredValue(TEXT("filter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("filter-text-color-dark"), RGB(255, 255, 255));
			int filterBackColor = !isDark ? getStoredValue(TEXT("filter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("filter-back-color-dark"), RGB(60, 60, 60));
			int currCellColor = !isDark ? getStoredValue(TEXT("current-cell-back-color"), RGB(70, 96, 166)) : getStoredValue(TEXT("current-cell-back-color-dark"), RGB(32, 62, 62));
			int selectionTextColor = !isDark ? getStoredValue(TEXT("selection-text-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("selection-text-color-dark"), RGB(220, 220, 220));
			int selectionBackColor = !isDark ? getStoredValue(TEXT("selection-back-color"), RGB(10, 36, 106)) : getStoredValue(TEXT("selection-back-color-dark"), RGB(72, 102, 102));
			int splitterColor = !isDark ? getStoredValue(TEXT("splitter-color"), GetSysColor(COLOR_BTNFACE)) : getStoredValue(TEXT("splitter-color-dark"), GetSysColor(COLOR_BTNFACE));			
			
			*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = textColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = backColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR2")) = backColor2;
			*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = filterTextColor;
			*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = filterBackColor;
			*(int*)GetProp(hWnd, TEXT("CURRENTCELLCOLOR")) = currCellColor;
			*(int*)GetProp(hWnd, TEXT("SELECTIONTEXTCOLOR")) = selectionTextColor;
			*(int*)GetProp(hWnd, TEXT("SELECTIONBACKCOLOR")) = selectionBackColor;
			*(int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")) = splitterColor;

			DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));			
			SetProp(hWnd, TEXT("BACKBRUSH"), CreateSolidBrush(backColor));
			SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));
			SetProp(hWnd, TEXT("SPLITTERBRUSH"), CreateSolidBrush(splitterColor));			

			ListView_SetTextColor(hGridWnd, textColor);			
			ListView_SetBkColor(hGridWnd, backColor);
			ListView_SetTextBkColor(hGridWnd, backColor);
			InvalidateRect(hWnd, NULL, TRUE);	
		}
		break;
		
		case WMU_HOT_KEYS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			if (wParam == VK_TAB) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				no += isCtrl ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isCtrl ? wnds[cnt - 1] : wnds[0]));
			}
			
			if (wParam == VK_F1) {
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/odbc-wlx/wiki"), 0, 0 , SW_SHOW);
				return TRUE;
			}
			
			if (wParam == 0x20 && isCtrl) { // Ctrl + Space
				SendMessage(hWnd, WMU_SHOW_COLUMNS, 0, 0);
				return TRUE;
			}
			
			if (wParam == VK_ESCAPE || wParam == VK_F11 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || (isCtrl && wParam == 0x46) || // Ctrl + F
				((wParam >= 0x31 && wParam <= 0x38) && !getStoredValue(TEXT("disable-num-keys"), 0) || // 1 - 8
				(wParam == 0x4E || wParam == 0x50) && !getStoredValue(TEXT("disable-np-keys"), 0)) && // N, P
				GetDlgCtrlID(GetFocus()) / 100 * 100 != IDC_HEADER_EDIT) { 
				SetFocus(GetParent(hWnd));		
				keybd_event(wParam, wParam, KEYEVENTF_EXTENDEDKEY, 0);

				return TRUE;
			}			
			
			return FALSE;					
		}
		break;
		
		case WMU_HOT_CHARS: {
			BOOL isCtrl = HIWORD(GetKeyState(VK_CONTROL));
			return !_istprint(wParam) && (
				wParam == VK_ESCAPE || wParam == VK_F11 || wParam == VK_F1 ||
				wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7) ||
				wParam == VK_TAB || wParam == VK_RETURN ||
				isCtrl && (wParam == 0x46 || wParam == 0x20);	
		}
		break;				
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && SendMessage(getMainWindow(hWnd), WMU_HOT_KEYS, wParam, lParam))
		return 0;

	// Prevent beep
	if (msg == WM_CHAR && SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
		return 0;	

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_CTLCOLOREDIT) {
		HWND hMainWnd = getMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));	
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HWND hMainWnd = getMainWindow(hWnd);
			BOOL isDark = *(int*)GetProp(hMainWnd, TEXT("DARKTHEME")); 

			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			LineTo(hDC, rc.right - 1, rc.bottom - 1);
			
			if (isDark) {
				DeleteObject(hPen);
				hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
				SelectObject(hDC, hPen);
				
				MoveToEx(hDC, 0, 0, 0);
				LineTo(hDC, 0, rc.bottom);
				MoveToEx(hDC, 0, rc.bottom - 1, 0);
				LineTo(hDC, rc.right, rc.bottom - 1);
				MoveToEx(hDC, 0, rc.bottom - 2, 0);
				LineTo(hDC, rc.right, rc.bottom - 2);
			}
			
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);
			
			return 0;
		}
		break;
		
		case WM_SETFOCUS: {
			SetProp(getMainWindow(hWnd), TEXT("LASTFOCUS"), hWnd);
		}
		break;	
		
		case WM_KEYDOWN: {
			HWND hMainWnd = getMainWindow(hWnd);
			if (wParam == VK_RETURN) {
				SendMessage(hMainWnd, WMU_UPDATE_CACHE, 0, 0);
				return 0;			
			}
			
			if (SendMessage(hMainWnd, WMU_HOT_KEYS, wParam, lParam))
				return 0;
		}
		break;
	
		// Prevent beep
		case WM_CHAR: {
			if (SendMessage(getMainWindow(hWnd), WMU_HOT_CHARS, wParam, lParam))
				return 0;	
		}
		break;		

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

HWND getMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;	
}

void setStoredValue(TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
}

TCHAR* getStoredString(TCHAR* name, TCHAR* defValue) { 
	TCHAR* buf = calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) && defValue)
		_tcsncpy(buf, defValue, 255);
	return buf;	
}

int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam) {
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd)) {
		int no = 0;
		HWND* wnds = (HWND*)lParam;
		while (wnds[no])
			no++;
		wnds[no] = hWnd;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* in) {
	TCHAR *out;
	if (!in || strlen(in) == 0) {
		out = (TCHAR*)calloc (1, sizeof (TCHAR));
	} else  {
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
		out = (TCHAR*)calloc (size, sizeof (TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, in, -1, out, size);
	}
	return out;
}

char* utf16to8(const TCHAR* in) {
	char* out;
	if (!in || _tcslen(in) == 0) {
		out = (char*)calloc (1, sizeof(char));
	} else  {
		int len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, 0, 0);
		out = (char*)calloc (len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, 0, 0);
	}
	return out;
}

int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;
		
	int res = -1;
	int tlen = _tcslen(text);
	int wlen = _tcslen(word);	
	if (!tlen || !wlen)
		return res;
	
	if (!isMatchCase) {
		TCHAR* ltext = calloc(tlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, text, tlen);
		text = _tcslwr(ltext);

		TCHAR* lword = calloc(wlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, word, wlen);
		word = _tcslwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res  == -1) && (pos <= tlen - wlen); pos++) 
			res = (pos == 0 || pos > 0 && !_istalnum(text[pos - 1])) && 
				!_istalnum(text[pos + wlen]) && 
				_tcsncmp(text + pos, word, wlen) == 0 ? pos : -1;
	} else {
		TCHAR* s = _tcsstr(text, word);
		res = s != NULL ? s - text : -1;
	}
	
	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res; 
}

TCHAR* extractUrl(TCHAR* data) {
	int len = data ? _tcslen(data) : 0;
	int start = len;
	int end = len;
	
	TCHAR* url = calloc(len + 10, sizeof(TCHAR));
	
	TCHAR* slashes = _tcsstr(data, TEXT("://"));
	if (slashes) {
		start = len - _tcslen(slashes);
		end = start + 3;
		for (; start > 0 && _istalpha(data[start - 1]); start--);
		for (; end < len && data[end] != TEXT(' ') && data[end] != TEXT('"') && data[end] != TEXT('\''); end++);
		_tcsncpy(url, data + start, end - start);
		
	} else if (_tcschr(data, TEXT('.'))) {
		_sntprintf(url, len + 10, TEXT("https://%ls"), data);
	}
	
	return url;
}

void setClipboardText(const TCHAR* text) {
	int len = (_tcslen(text) + 1) * sizeof(TCHAR);
	HGLOBAL hMem =  GlobalAlloc(GMEM_MOVEABLE, len);
	memcpy(GlobalLock(hMem), text, len);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

BOOL isNumber(TCHAR* val) {
	int len = _tcslen(val);
	BOOL res = TRUE;
	int pCount = 0;
	for (int i = 0; res && i < len; i++) {
		pCount += val[i] == TEXT('.');
		res = _istdigit(val[i]) || val[i] == TEXT('.');
	}
	return res && pCount < 2;
}

int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt) {
	int colNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM | LVCF_FMT;
	lvc.iSubItem = colNo;
	lvc.pszText = colName;
	lvc.cchTextMax = _tcslen(colName) + 1;
	lvc.cx = 100;
	lvc.fmt = fmt;
	return ListView_InsertColumn(hListWnd, colNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax) {
	if (i < 0)
		return FALSE;

	TCHAR buf[cchTextMax];

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = buf;
	hdi.cchTextMax = cchTextMax;
	int rc = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, buf, cchTextMax);
	return rc;
}

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState) {
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = fState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}
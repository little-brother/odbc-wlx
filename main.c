#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
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
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 4
#define WMU_RESET_CACHE        WM_USER + 5
#define WMU_SET_FONT           WM_USER + 6
#define WMU_ERROR_MESSAGE      WM_USER + 7

#define IDC_MAIN               100
#define IDC_TABLELIST          101
#define IDC_DATAGRID           102
#define IDC_STATUSBAR          103
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROW           5001

#define IDH_EXIT               6000
#define IDH_NEXT               6001
#define IDH_PREV               6002

#define SB_TABLE_COUNT         0
#define SB_VIEW_COUNT          1
#define SB_TYPE                2
#define SB_ROW_COUNT           3
#define SB_CURRENT_ROW         4
#define SB_ERROR               5

#define MAX_TEXT_LENGTH        32000
#define MAX_DATA_LENGTH        32000
#define MAX_COLUMN_LENGTH      2000
#define MAX_ERROR_LENGTH       2000

#define ODBC_ACCESS            1
#define ODBC_EXCEL             2
#define ODBC_EXCELX            3
#define ODBC_CSV               4

#define APP_NAME               TEXT("odbc-wlx")

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
void setClipboardText(const TCHAR* text);
BOOL isNumber(TCHAR* val);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName, int fmt);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	snprintf(DetectString, maxlen, "MULTIMEDIA & (ext=\"MDB\" | ext=\"XLS\" | ext=\"XLSX\" | ext=\"XLSB\" | ext=\"CSV\" | ext=\"DSN\")");
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* filepath = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, filepath, size);

	int odbcType = ODBC_CSV;

	TCHAR* fileext = _tcsrchr(filepath, TEXT('.'));
	_tcslwr(fileext);

	int dlen = _tcslen(filepath);
	while (dlen > 0 && filepath[dlen - 1] != TEXT('/') && filepath[dlen - 1] != TEXT('\\'))
		dlen--;

	TCHAR connectionString[MAX_TEXT_LENGTH] = {0};
	TCHAR connectionString2[MAX_TEXT_LENGTH] = {0};
	if (_tcscmp(fileext, TEXT(".mdb")) == 0 || _tcscmp(fileext, TEXT(".accdb")) == 0) {
		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Access Driver (*.mdb)};Dbq=%ls;Uid=Admin;Pwd=;ReadOnly=0;"), filepath);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=%ls;Uid=Admin;Pwd=;ReadOnly=0;"), filepath);
		odbcType = ODBC_ACCESS;
	} else if (_tcscmp(fileext, TEXT(".xls")) == 0) {
		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls)};Dbq=%ls;ReadOnly=0;"), filepath);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=%ls;ReadOnly=0;"), filepath);
		odbcType = ODBC_EXCEL;
	} else if (_tcscmp(fileext, TEXT(".xlsx")) == 0 || _tcscmp(fileext, TEXT(".xlsb")) == 0) {
		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=%ls;ReadOnly=0;"), filepath);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};Dbq=%ls;ReadOnly=0;"), filepath);
		odbcType = ODBC_EXCELX;
	} else if (_tcscmp(fileext, TEXT(".csv")) == 0) {
		TCHAR dir[MAX_PATH] = {0};
		_tcsncpy(dir, filepath, dlen - 1);

		_sntprintf(connectionString, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=%ls;Extensions=asc,csv,tab,txt;ReadOnly=0;"), dir);
		_sntprintf(connectionString2, MAX_TEXT_LENGTH, TEXT("Driver={Microsoft Access Text Driver (*.txt, *.csv)};Dbq=%ls; Extensions=asc,csv,tab,txt;ReadOnly=0;"), dir);
		odbcType = ODBC_CSV;
	} else if (_tcscmp(fileext, TEXT(".dsn")) == 0) {
		TCHAR buf[32000];

		int len = GetPrivateProfileString(TEXT("ODBC"), NULL, NULL, buf, 32000, filepath);
		int start = 0;
		for (int i = 0; i < len; i++) {
			if (buf[i] != 0)
				continue;

			TCHAR key[i - start + 1];
			_tcsncpy(key, buf + start, i - start + 1);
			TCHAR value[1024];
			GetPrivateProfileString(TEXT("ODBC"), key, NULL, value, 1024, filepath);
			TCHAR pair[2000];
			BOOL isQ = _tcschr(value, TEXT(' ')) != 0;
			_sntprintf(pair, 2000, TEXT("%ls=%ls%ls%ls;"), key, isQ ? TEXT("{") : TEXT(""), value, isQ ? TEXT("}") : TEXT(""));
			_tcscat(connectionString, pair);

			start = i + 1;

			_tcslwr(key);
			_tcslwr(value);
			if (_tcscmp(key, TEXT("driver")) == 0)
				odbcType = _tcsstr(value, TEXT("*.csv")) ? ODBC_CSV : _tcsstr(value, TEXT("*.mdb")) ? ODBC_ACCESS : ODBC_EXCEL;
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
		return 0;
	}

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES;
	InitCommonControlsEx(&icex);

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindow(WC_STATIC, TEXT("odbc-wlx"), WS_CHILD | WS_VISIBLE | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TABLENAME"), calloc(1024, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("WHERE"), calloc(MAX_TEXT_LENGTH, sizeof(TCHAR)));
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("DBENV"), hEnv);
	SetProp(hMainWnd, TEXT("DB"), hConn);
	SetProp(hMainWnd, TEXT("COLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ODBCTYPE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERWIDTH"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("GRAYBRUSH"), CreateSolidBrush(GetSysColor(COLOR_BTNFACE)));

	*(int*)GetProp(hMainWnd, TEXT("SPLITTERWIDTH")) = getStoredValue(TEXT("splitter-width"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);
	*(int*)GetProp(hMainWnd, TEXT("ODBCTYPE")) = odbcType;	

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	int sizes[6] = {75, 150, 200, 400, 500, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 6, (LPARAM)&sizes);

	HWND hListWnd = CreateWindow(TEXT("LISTBOX"), NULL, WS_CHILD | WS_VISIBLE | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | WS_VSCROLL | WS_TABSTOP,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_TABLELIST, GetModuleHandle(0), NULL);

	HWND hDataWnd = CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hMainWnd, (HMENU)IDC_DATAGRID, GetModuleHandle(0), NULL);
	ListView_SetExtendedListViewStyle(hDataWnd, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);

	HWND hHeader = ListView_GetHeader(hDataWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);

	HMENU hDataMenu = CreatePopupMenu();
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_ROW, TEXT("Copy row"));
	SetProp(hMainWnd, TEXT("DATAMENU"), hDataMenu);

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

		if (odbcType == ODBC_ACCESS && _tcscmp(tblType, TEXT("system table")) == 0)
			continue;

		int pos = ListBox_AddString(hListWnd, tblName);
		BOOL isTable = _tcsstr(tblType, TEXT("view")) == 0;
		SendMessage(hListWnd, LB_SETITEMDATA, pos, isTable);

		tCount += isTable;
		vCount += !isTable;
	}
	SQLCloseCursor(hStmt);
	SQLFreeHandle(SQL_HANDLE_STMT, hStmt);

	TCHAR buf[255];
	_sntprintf(buf, 255, TEXT(" Tables: %i"), tCount);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_TABLE_COUNT, (LPARAM)buf);
	_sntprintf(buf, 255, TEXT(" Views: %i"), vCount);
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VIEW_COUNT, (LPARAM)buf);
	
	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WM_SIZE, 0, 0);
	
	ListBox_SetCurSel(hListWnd, _tcscmp(fileext, TEXT(".csv")) == 0 ? ListBox_FindStringExact(hListWnd, -1, filepath + dlen) : 0);
	SendMessage(hMainWnd, WMU_UPDATE_GRID, 0, 0);

	RegisterHotKey(hMainWnd, IDH_EXIT, 0, VK_ESCAPE);
	RegisterHotKey(hMainWnd, IDH_NEXT, 0, VK_TAB);
	RegisterHotKey(hMainWnd, IDH_PREV, MOD_CONTROL, VK_TAB);
	SetFocus(hListWnd);	

	return hMainWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-width"), *(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	
	SQLHANDLE hEnv = (SQLHANDLE)GetProp(hWnd, TEXT("DBENV"));
	SQLHANDLE hConn = (SQLHANDLE)GetProp(hWnd, TEXT("DB"));
	SQLDisconnect(hConn);
	SQLFreeHandle(SQL_HANDLE_DBC, hConn);
	SQLFreeHandle(SQL_HANDLE_ENV, hEnv);

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((TCHAR*)GetProp(hWnd, TEXT("TABLENAME")));
	free((TCHAR*)GetProp(hWnd, TEXT("WHERE")));
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("SPLITTERWIDTH")));
	free((int*)GetProp(hWnd, TEXT("FONTSIZE")));	
	free((int*)GetProp(hWnd, TEXT("COLNO")));
	free((int*)GetProp(hWnd, TEXT("ODBCTYPE")));

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("GRAYBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("DATAMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("DBENV"));
	RemoveProp(hWnd, TEXT("DB"));
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("TABLENAME8"));
	RemoveProp(hWnd, TEXT("WHERE8"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("SPLITTERWIDTH"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));	
	RemoveProp(hWnd, TEXT("COLNO"));
	RemoveProp(hWnd, TEXT("ODBCTYPE"));

	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("GRAYBRUSH"));
	RemoveProp(hWnd, TEXT("DATAMENU"));

	DestroyWindow(hWnd);
	return;
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_HOTKEY: {
			WPARAM id = wParam;
			if (id == IDH_EXIT)
				SendMessage(GetParent(hWnd), WM_CLOSE, 0, 0);

			if (id == IDH_NEXT || id == IDH_PREV) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				BOOL isBackward = id == IDH_PREV;
				no += isBackward ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isBackward ? wnds[cnt - 1] : wnds[0]));
			}
		}
		break;

		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int splitterW = *(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH"));
			GetClientRect(hWnd, &rc);
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			SetWindowPos(hListWnd, 0, 0, 0, splitterW, rc.bottom - statusH, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hDataWnd, 0, splitterW + 5, 0, rc.right - splitterW - 5, rc.bottom - statusH, SWP_NOZORDER);
		}
		break;

		case WM_PAINT: {
			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			rc.left = *(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH"));
			rc.right = rc.left + 5;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("GRAYBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;

		case WM_LBUTTONDOWN: {
			SetProp(hWnd, TEXT("ISMOUSEDOWN"), (HANDLE)1);
			SetCapture(hWnd);
			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
		}
		break;

		case WM_MOUSEMOVE: {
			if (wParam != MK_LBUTTON || !GetProp(hWnd, TEXT("ISMOUSEDOWN")))
				return 0;

			DWORD x = GET_X_LPARAM(lParam);
			if (x > 0 && x < 32000)
				*(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH")) = x;
			SendMessage(hWnd, WM_SIZE, 0, 0);
		}
		break;
		
		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
			}
		}
		break;		

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDC_TABLELIST && HIWORD(wParam) == LBN_SELCHANGE)
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);

			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROW) {
				HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
				HWND hHeader = ListView_GetHeader(hDataWnd);
				int rowNo = ListView_GetNextItem(hDataWnd, -1, LVNI_SELECTED);
				if (rowNo != -1) {
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

					int colCount = Header_GetItemCount(hHeader);

					int startNo = cmd == IDM_COPY_CELL ? *(int*)GetProp(hWnd, TEXT("COLNO")) : 0;
					int endNo = cmd == IDM_COPY_CELL ? startNo + 1 : colCount;
					if (startNo > colCount || endNo > colCount)
						return 0;

					int len = 0;
					for (int colNo = startNo; colNo < endNo; colNo++)
						len += _tcslen(cache[rowNo][colNo]) + 1 /* column delimiter: TAB */;

					TCHAR buf[len + 1];
					buf[0] = 0;
					for (int colNo = startNo; colNo < endNo; colNo++) {
						_tcscat(buf, cache[rowNo][colNo]);
						if (colNo != endNo - 1)
							_tcscat(buf, TEXT("\t"));
					}

					setClipboardText(buf);
				}
			}
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_DATAGRID && pHdr->code == LVN_GETDISPINFO) {
				LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
				LV_ITEM* pItem= &(pDispInfo)->item;
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

				if(pItem->mask & LVIF_TEXT)
					_tcsncpy(pItem->pszText, cache[pItem->iItem][pItem->iSubItem], pItem->cchTextMax);
			}

			if (pHdr->idFrom == IDC_DATAGRID && pHdr->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* pLV = (NMLISTVIEW*)lParam;
				int colNo = pLV->iSubItem + 1;
				int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
				int orderBy = *pOrderBy;
				*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
				SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
			}

			if (pHdr->idFrom == IDC_DATAGRID && pHdr->code == (DWORD)NM_RCLICK) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;

				POINT p;
				GetCursorPos(&p);
				*(int*)GetProp(hWnd, TEXT("COLNO")) = ia->iSubItem;
				TrackPopupMenu(GetProp(hWnd, TEXT("DATAMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}

			if (pHdr->idFrom == IDC_DATAGRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);

				TCHAR buf[255] = {0};
				int pos = ListView_GetNextItem(pHdr->hwndFrom, -1, LVNI_SELECTED);
				if (pos != -1)
					_sntprintf(buf, 255, TEXT(" %i"), pos + 1);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_ROW, (LPARAM)buf);
			}

			if (pHdr->idFrom == IDC_DATAGRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43 && GetKeyState(VK_CONTROL)) // Ctrl + C
					SendMessage(hWnd, WM_COMMAND, IDM_COPY_ROW, 0);
			}

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(hWnd, IDC_DATAGRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_GRID: {
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SQLHANDLE hConn = (SQLHANDLE)GetProp(hWnd, TEXT("DB"));

			SendMessage(hDataWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hDataWnd);

			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			ListView_SetItemCount(hDataWnd, 0);

			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++) 
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hDataWnd, colCount - colNo - 1);

			TCHAR* tablename = (TCHAR*)GetProp(hWnd, TEXT("TABLENAME"));
			int pos = ListBox_GetCurSel(hListWnd);
			ListBox_GetText(hListWnd, pos, tablename);

			TCHAR buf[255];
			int type = SendMessage(hListWnd, LB_GETITEMDATA, pos, 0);
			_sntprintf(buf, 255, type ? TEXT("  TABLE"): TEXT("   VIEW"));
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TYPE, (LPARAM)buf);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_ROW, (LPARAM)0);

			SQLHANDLE hStmt = 0;
			SQLAllocHandle(SQL_HANDLE_STMT, hConn, &hStmt);
			TCHAR query[1024 + _tcslen(tablename)];
			_sntprintf(query, 1024 + _tcslen(tablename), TEXT("select * from \"%ls\" where 1 = 2"), tablename);
			if (SQL_SUCCESS == SQLExecDirect(hStmt, query, SQL_NTS)) {
				SQLSMALLINT colCount = 0;
				SQLNumResultCols(hStmt, &colCount);

				for (int colNo = 0; colNo < colCount; colNo++) {
					SQLWCHAR colName[MAX_COLUMN_LENGTH];
					SQLSMALLINT colType = 0;
					SQLDescribeCol(hStmt, colNo + 1, colName, MAX_COLUMN_LENGTH, 0, &colType, 0, 0, 0);

					int fmt = colType == SQL_DECIMAL || colType == SQL_NUMERIC || colType == SQL_REAL || colType == SQL_FLOAT || colType == SQL_DOUBLE ||
						colType == SQL_SMALLINT || colType == SQL_INTEGER || colType == SQL_BIT || colType == SQL_TINYINT || colType == SQL_BIGINT ?
						LVCFMT_RIGHT :
						LVCFMT_LEFT;
						
					ListView_AddColumn(hDataWnd, colName, fmt);	
				}

				for (int colNo = 0; colNo < colCount; colNo++) {
					RECT rc;
					Header_GetItemRect(hHeader, colNo, &rc);
					HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, ES_CENTER | ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD | WS_BORDER | WS_TABSTOP, 0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
					SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
					SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
				}
			} else {
				SendMessage(hWnd, WMU_ERROR_MESSAGE, (WPARAM)hStmt, 0);
			}
			SQLCloseCursor(hStmt);
			SQLFreeHandle(SQL_HANDLE_STMT, hStmt);

			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;
			SendMessage(hWnd, WMU_UPDATE_CACHE, 0, 0);
			SendMessage(hDataWnd, WM_SETREDRAW, TRUE, 0);

			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_CACHE: {
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hDataWnd);
			int colCount = Header_GetItemCount(hHeader);
			SQLHANDLE hConn = (SQLHANDLE)GetProp(hWnd, TEXT("DB"));
			TCHAR* tablename = (TCHAR*)GetProp(hWnd, TEXT("TABLENAME"));
			TCHAR* where = (TCHAR*)GetProp(hWnd, TEXT("WHERE"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int orderBy = *(int*)GetProp(hWnd, TEXT("ORDERBY"));

			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			ListView_SetItemCount(hDataWnd, 0);
			SendMessage(hWnd, WMU_ERROR_MESSAGE, 0, 0);

			_sntprintf(where, MAX_TEXT_LENGTH, TEXT("where (1 = 1)"));
			for (int colNo = 0; colNo < colCount; colNo++) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				int len = GetWindowTextLength(hEdit);
				if (len > 0) {
					TCHAR colName[256] = {0};
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
						val[0] == TEXT('!') ? TEXT("\" not like '%%' & ? & '%%'") :
						val[0] == TEXT('>') ? TEXT("\" > %ls") :
						val[0] == TEXT('<') ? TEXT("\" < %ls") :
						TEXT("\" like '%%' & %ls & '%%'"), isNumber(val + hasPrefix) ? val + hasPrefix : qval);

					_tcscat(where, cond);
				}
			}

			int rowCount = -1;

			SQLHANDLE hStmt = 0;
			SQLAllocHandle(SQL_HANDLE_STMT, hConn, &hStmt);
			int len = 1024 + _tcslen(tablename) + _tcslen(where);
			TCHAR* query = calloc(len + 1, sizeof(TCHAR));
			_sntprintf(query, len, TEXT("select count(*) from \"%ls\" %ls"), tablename, where);
			if((SQL_SUCCESS == SQLExecDirect(hStmt, query, SQL_NTS)) && (SQL_SUCCESS == SQLFetch(hStmt))) {
				SQLLEN res = 0;
				SQLGetData(hStmt, 1, SQL_C_LONG, &rowCount, sizeof(int), &res);
			}
			free(query);

			if (rowCount == -1) {
				SendMessage(hWnd, WMU_ERROR_MESSAGE, (WPARAM)hStmt, 0);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)TEXT("N/A"));
			}
			SQLCloseCursor(hStmt);
			SQLFreeHandle(SQL_HANDLE_STMT, hStmt);

			int rowLimit = getStoredValue(TEXT("max-row-count"), 0);
			if (rowLimit > 0 && rowCount > rowLimit && rowCount != -1) 
				rowCount = rowLimit;

			if (_tcscmp(where, TEXT("where (1 = 1)")) == 0)
				*pTotalRowCount = rowCount;
			*pRowCount = rowCount;

			if (rowCount == -1)
				return 0;

			TCHAR buf[1024];
			_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), rowCount, *pTotalRowCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);

			ListView_SetItemCount(hDataWnd, rowCount);

			if (rowCount == 0)
				return 0;

			SetProp(hWnd, TEXT("CACHE"), calloc(rowCount, sizeof(TCHAR*)));
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

			SQLAllocHandle(SQL_HANDLE_STMT, hConn, &hStmt);
			len = 1024 + _tcslen(tablename) + _tcslen(where);
			query = calloc(len + 1, sizeof(TCHAR));
			TCHAR orderBy16[32] = {0};
			if (orderBy > 0)
				_sntprintf(orderBy16, 32, TEXT("order by %i"), orderBy);
			if (orderBy < 0)
				_sntprintf(orderBy16, 32, TEXT("order by %i desc"), -orderBy);
			_sntprintf(query, len, TEXT("select * from \"%ls\" %ls %ls"), tablename, where, orderBy16);
			
			int rowNo = 0;
			if(SQL_SUCCESS == SQLExecDirect(hStmt, query, SQL_NTS)) {
				while (SQLFetch(hStmt) == SQL_SUCCESS && rowNo < rowCount) {						
					cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));

					for (int colNo = 0; colNo < colCount; colNo++) {
						SQLLEN res = 0;
						SQLWCHAR val[MAX_DATA_LENGTH];
						SQLGetData(hStmt, colNo + 1, SQL_C_TCHAR, val, MAX_DATA_LENGTH * sizeof(TCHAR), &res);
						cache[rowNo][colNo] = calloc(res + 2, sizeof(TCHAR)); // res = -1 for NULL
						if (res > 0)
							_tcsncpy(cache[rowNo][colNo], val, res);
					}

					rowNo++;
				}
			}
			SQLCloseCursor(hStmt);
			SQLFreeHandle(SQL_HANDLE_STMT, hStmt);

			free(query);
		}
		break;

		case WMU_UPDATE_FILTER_SIZE: {
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			HWND hHeader = ListView_GetHeader(hDataWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);
			for (int colNo = 0; colNo < colCount; colNo++) {
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left - (colNo > 0), h2, rc.right - rc.left + 1, h2 + 1, SWP_NOZORDER);							
			}
		}
		break;

		case WMU_AUTO_COLUMN_SIZE: {
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			SendMessage(hDataWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hDataWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hDataWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hDataWnd, 0) < 100)
				ListView_SetColumnWidth(hDataWnd, 0, 100);
				
			int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hDataWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hDataWnd, colNo, maxWidth);
				}
			}

			// Fix last column				
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hDataWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[1024];
				Header_GetItemText(hHeader, colNo, name16, 1024);
				RECT rc;
				HDC hDC = GetDC(hHeader);
				DrawText(hDC, name16, _tcslen(name16), &rc, DT_NOCLIP | DT_CALCRECT);
				ReleaseDC(hHeader, hDC);

				int w = rc.right - rc.left + 10;
				if (ListView_GetColumnWidth(hDataWnd, colNo) < w)
					ListView_SetColumnWidth(hDataWnd, colNo, w);
			}

			SendMessage(hDataWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hDataWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_RESET_CACHE: {
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));

			int colCount = Header_GetItemCount(ListView_GetHeader(hDataWnd));
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

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, TEXT("Arial"));
			HWND hListWnd = GetDlgItem(hWnd, IDC_TABLELIST);
			HWND hDataWnd = GetDlgItem(hWnd, IDC_DATAGRID);
			SendMessage(hListWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hDataWnd, WM_SETFONT, (LPARAM)hFont, TRUE);

			HWND hHeader = ListView_GetHeader(hDataWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

			SetProp(hWnd, TEXT("FONT"), hFont);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;		

		// wParam = hStmt
		case WMU_ERROR_MESSAGE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SQLHANDLE hStmt = (SQLHANDLE)wParam;
			if (hStmt) {
				SQLWCHAR err[MAX_ERROR_LENGTH + 1];
				SQLWCHAR code[6];
				SQLGetDiagRec(SQL_HANDLE_STMT, hStmt, 1, code, NULL, err, MAX_ERROR_LENGTH, NULL);
				TCHAR msg[MAX_ERROR_LENGTH + 100];
				_sntprintf(msg, MAX_ERROR_LENGTH + 100, TEXT("Error (%ls): %ls"), code, err);

				SendMessage(hStatusWnd, SB_SETTEXT, SB_ERROR, (LPARAM)msg);
			} else {
				SendMessage(hStatusWnd, SB_SETTEXT, SB_ERROR, (LPARAM)TEXT(""));
			}
		}
		break;
	}
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		// Win10+ fix: draw an upper border
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetWindowRect(hWnd, &rc);
			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;

		// Prevent beep
		case WM_CHAR: {
			if (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB)
				return 0;
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB) {
				if (wParam == VK_RETURN) {
					HWND hHeader = GetParent(hWnd);
					HWND hDataWnd = GetParent(hHeader);
					HWND hMainWnd = GetParent(hDataWnd);
					SendMessage(hMainWnd, WMU_UPDATE_CACHE, 0, 0);
				}

				return 0;
			}
		}
		break;

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
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
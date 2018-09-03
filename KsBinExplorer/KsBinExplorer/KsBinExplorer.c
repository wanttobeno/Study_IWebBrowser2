/****************************************************************************
 *                                                                          *
 * File    : KsBinExplorer.c                                                         *
 *                                                                          *
 * Purpose : Generic dialog based Win32 application.                        *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

/* 
 * Either define WIN32_LEAN_AND_MEAN, or one or more of NOCRYPT,
 * NOSERVICE, NOMCX and NOIME, to decrease compile time (if you
 * don't need these defines -- see windows.h).
 */

#define WIN32_LEAN_AND_MEAN
/* #define NOCRYPT */
/* #define NOSERVICE */
/* #define NOMCX */
/* #define NOIME */
#include "stdafx.h"
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <tchar.h>
#include <stdio.h>
#include "resource.h"
#include "TabCtrl.c"

#include "globalVariable.h"

#include "CHTMLTabCtrl.c"
#pragma comment(lib,"comctl32.lib") 
#pragma comment(lib,"version.lib") 


/** Macroes *****************************************************************/

// #define Refresh(A) RedrawWindow(A,NULL,NULL,RDW_ERASE|RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW)



/** Prototypes **************************************************************/


// URl框的消息函数，子类化出来接受回车键
LRESULT CALLBACK FormMain_EditProc(HWND hWnd,UINT uMsg,WPARAM wParam,LPARAM lParam)
{
    //这里就是一直希望可以自己处理的WM_CHAR消息了，0-9和BackSpace键放行
    if (uMsg==WM_CHAR)
    {
        if( 0xd ==wParam )
        {
			cmdClickRunUrl_Click (g_frmMain);
        }
		
    }
    //调用原来的默认消息处理函数，和DefWindowProc意思一样
    return CallWindowProc(wpFormMain_OrigEditProc, hWnd, uMsg,wParam, lParam); 
}

static void InitHandles (HWND hwndParent)
{
	static BOOL initialized=FALSE;
	if(!initialized)
	{
		if (g_frmMain != hwndParent)
			g_frmMain = hwndParent;

		//Get a handle to the Menu Bar
		if (mnuMain != GetMenu(hwndParent))
			mnuMain = GetMenu(hwndParent);

		initialized=TRUE;
	}
}

/****************************************************************************
 *                                                                          *
 * Function: DoEvents					                                    *
 *                                                                          *
 * Purpose: DoEvents is a statement that yields execution of the current	*
 * thread so that the operating system can process other events.			*
 * This function cleans out the message loop and executes any other pending	*
 * business.																*
 *																			*
 ****************************************************************************/

void DoEvents (void)
{
	MSG Msg;
	while (PeekMessage(&Msg,NULL,0,0,PM_REMOVE))
	{
		TranslateMessage(&Msg);
		DispatchMessage(&Msg);
	}
}

/****************************************************************************
 *                                                                          *
 * Function: GetVersionInfo                                                 *
 *                                                                          *
 * Purpose : Returns the Info associated with a version info entry string.  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Adapted from luetz                                   *
 *					   www.codeproject.com/cpp/GetLocalVersionInfos.asp		*
 *                                                                          *
 * Usage : GetVersionInfo(FieldDescriptionString)                           *
 *                                                                          *
 *		   Where FieldDescriptionString = "Comments"						*
 *										  "CompanyName"						*
 *										  "FileDescription"					*
 *										  "InternalName"					*
 *										  "LegalCopyright"					*
 *										  "LegalTrademarks"					*
 *										  "OriginalFilename"				*
 *										  "PrivateBuild"					*
 *										  "ProductName"						*
 *										  "ProductVersion"					*
 *										  "SpecialBuild"					*
 *                                                                          *
 ****************************************************************************/

LPWSTR GetVersionInfo(LPWSTR csEntry)
{
	LPWSTR csRet;
	static WCHAR errStr[MAX_PATH];

	HRSRC hVersion = FindResource(GetModuleHandle(NULL), 
	MAKEINTRESOURCE(VS_VERSION_INFO), RT_VERSION );
	if (hVersion!=NULL)
	{
		HGLOBAL hGlobal = LoadResource(GetModuleHandle(NULL), hVersion );
 
		if (hGlobal!=NULL)  
		{  
			LPVOID versionInfo=LockResource(hGlobal);
  
			if (versionInfo!=NULL)
			{
				DWORD vLen,langD;
				BOOL retVal;    
				LPVOID retbuf=NULL;
				static WCHAR fileEntry[256];

				wsprintf(fileEntry, L"\\VarFileInfo\\Translation");
				retVal = VerQueryValue(versionInfo,fileEntry,&retbuf,(UINT *)&vLen);
				if (retVal && vLen==4) 
				{
					memcpy(&langD,retbuf,4);            
					wsprintf(fileEntry, L"\\StringFileInfo\\%02X%02X%02X%02X\\%s",
					(langD & 0xff00)>>8,langD & 0xff,(langD & 0xff000000)>>24, 
					(langD & 0xff0000)>>16, csEntry);            
				}
				else
				{ 
				wsprintf(fileEntry, L"\\StringFileInfo\\%04X04B0\\%s", 
				GetUserDefaultLangID(), csEntry);
				}
				if (retVal==VerQueryValue(versionInfo,fileEntry,&retbuf,(UINT *)&vLen))
				{ 
					csRet=(WCHAR*)retbuf;
				}
				else
				{
					wsprintf(errStr, L"%s Not Available",csEntry);
					csRet = errStr;
				}
			}
		}
		UnlockResource(hGlobal);  
		FreeResource(hGlobal);  
	}
	return csRet;
}
LRESULT FormMain_OnChangeTabLabel()
{
	HRESULT			hr = -1;
	IWebBrowser2	*pBrowserApp;
	IOleObject		*pBrowserObject;
	BSTR			lpbResultUrl = NULL;
	BSTR			lpbResultTitle = NULL;
	WCHAR			wszMyResultTitle[7] = {0};
	int				i = 1;
	int				iSel = TabCtrl_GetCurSel( g_MainWebTabCtrl.hTab );
	TCITEMW			tie;
	tie.mask = TCIF_TEXT | TCIF_IMAGE;
	tie.iImage = -1;
	if ( !g_MainWebTabCtrl.hVisiblePage )
	{
		return S_OK;
	}
	pBrowserObject = *((IOleObject **)GetWindowLong( g_MainWebTabCtrl.hVisiblePage, GWL_USERDATA ));
	hr = pBrowserObject->lpVtbl->QueryInterface(pBrowserObject, &IID_IWebBrowser2, (void**) &pBrowserApp);
	if ( SUCCEEDED(hr) )
	{		
		// ignore js error
		pBrowserApp->lpVtbl->put_Silent(pBrowserApp,VARIANT_TRUE);

		hr = pBrowserApp->lpVtbl->get_LocationName(pBrowserApp, &lpbResultTitle);
		if ( S_OK == (hr) )
		{

			memcpy( wszMyResultTitle, lpbResultTitle, 5*sizeof(WCHAR) );
			wszMyResultTitle[6] = L'…';
			tie.pszText = (LPWSTR)wszMyResultTitle;
			TabCtrl_SetItem(g_MainWebTabCtrl.hTab, iSel, &tie);
		}
		else{
			SetWindowTextW( g_frmMain , L"Fail !");
		}
	}
	return S_OK;
}
LRESULT FormMain_OnChangeUrlAndTitle()
{
	HRESULT			hr = -1;
	IWebBrowser2	*pBrowserApp;
	IOleObject		*pBrowserObject;
	BSTR			lpbResultUrl = NULL;
	BSTR			lpbResultTitle = NULL;
	WCHAR			wszMyResultTitle[7] = {0};
	int i = 1;
	int iSel		= TabCtrl_GetCurSel( g_MainWebTabCtrl.hTab );
	TCITEMW		tie;
	tie.mask = TCIF_TEXT | TCIF_IMAGE;
	tie.iImage = -1;
	if ( !g_MainWebTabCtrl.hVisiblePage )
	{
		return S_OK;
	}
	
	pBrowserObject = *((IOleObject **)GetWindowLong( g_MainWebTabCtrl.hVisiblePage, GWL_USERDATA ));
	hr = pBrowserObject->lpVtbl->QueryInterface(pBrowserObject, &IID_IWebBrowser2, (void**) &pBrowserApp);
	if ( SUCCEEDED(hr) )
	{			
		
		hr = pBrowserApp->lpVtbl->get_LocationURL(pBrowserApp, &lpbResultUrl);	
		if ( S_OK == (hr) )
		{
			SetWindowTextW( GetDlgItem( g_frmMain, IDC_EDIT_URL) , (LPCWSTR)lpbResultUrl);
		}
		else{
			SetWindowTextW( GetDlgItem( g_frmMain, IDC_EDIT_URL) , L"Fail !");
		}

		hr = pBrowserApp->lpVtbl->get_LocationName(pBrowserApp, &lpbResultTitle);
		if ( S_OK == (hr) )
		{
			SetWindowTextW( g_frmMain , (LPWSTR)lpbResultTitle);
		}
		else{
			SetWindowTextW( g_frmMain , L"Fail !");
		}
	}
	return S_OK;

}

static LRESULT FormMain_OnNotify(HWND hwnd, INT id, LPNMHDR pnm)
{
	switch (pnm->code)
	{
	case TCN_KEYDOWN:	
		 //TabCtrl_OnKeyDown((LPARAM)pnm);
		 // fall through to call TabCtrl_OnSelChanged() on each keydown 
		 break;
		
	case TCN_SELCHANGE:
		 FormMain_OnChangeUrlAndTitle();
		 break;
	}
    // 交给底下的控件 （tabcontrol）处理
    switch(id)
	{
// 		case TAB_CONTROL_1:
// 			return TabCtrl_1.Notify(pnm);

		case MAIN_WEB_TAB_CONTROL:
			return g_MainWebTabCtrl.Notify(pnm);

		//
		// TODO: Add other control id case statements here. . .
		//
	}
	return S_OK;
}

void FormMain_OnClose(HWND hwnd)
{
	PostQuitMessage(0);// turn off message loop

	CHtmlTabControl_Destroy( &g_MainWebTabCtrl );
	EndDialog(hwnd, 0);
}

/****************************************************************************
 *                                                                          *
 * Functions: FormMain_OnCommand related event code                         *
 *                                                                          *
 * Purpose : Handle WM_COMMAND messages: this is the heart of the app.		*
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/
void mnuAbout_Click (HWND hwnd)
{
	WCHAR buf[MAX_PATH];
	wsprintf(buf, L"Version = %ws\n",GetVersionInfo(L"ProductVersion"));
	MessageBox (hwnd,buf,GetVersionInfo(L"ProductName"),MB_OK);
// 	PostQuitMessage(0);// turn off message loop
// 	CHtmlTabControl_Destroy(&gMainWebTabCtrl);
// 	EndDialog(hwnd, 0);
}

void cmdClickWeb_DoAction (HWND hwnd, int iAction)
{
	CHtmlTabControl_DoAction(  hwnd,  iAction);
}

void cmdClickRunUrl_Click (HWND hwnd)
{
	PWCHAR	wszUrl;
	HWND	hUrl	=	GetDlgItem( hwnd,IDC_EDIT_URL );
	int		iUrlLen  =	GetWindowTextLengthW( hUrl ) + 1;
	if ( !iUrlLen)
	{
		return;
	}
			
	wszUrl	=	(PWCHAR)malloc( sizeof(WCHAR)*(iUrlLen+1) );
	if ( !wszUrl )
	{
		MessageBoxW(0, L"警告", L"内存不足！", 0);
		return;
	}
	GetWindowTextW(hUrl, wszUrl, iUrlLen);
	wszUrl[iUrlLen] = '\0';
	DisplayHTMLPage(CHtml_m_lptc->hVisiblePage, wszUrl );
}

void cmdClickWeb_Close (HWND hwnd)
{
	CHtmlTabControl_Close();
}

void FormMain_OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
	switch (id)
	{
		case MNU_ABOUT:
			mnuAbout_Click (hwnd);
		break;

		case MNU_QUIT:
			FormMain_OnClose (hwnd);
		break;

		case IDC_BUTTON_RUNURL:
			cmdClickRunUrl_Click (hwnd);
		break;

		case IDC_BUTTON_GOFORWARD:
		case IDC_BUTTON_GOBACK:
		case IDC_BUTTON_FLUSH:
		case IDC_BUTTON_STOP:
			cmdClickWeb_DoAction( hwnd, id );
		break;

		case IDC_BUTTON_CLOSE:
			cmdClickWeb_Close (hwnd);
		break;
		

		default: break;
	}
}

void FormMain_OnSize(HWND hwnd, UINT state, int cx, int cy)
{
	RECT  rc;
	int i=0;
	int iHight = 40;
   	GetClientRect(hwnd,&rc);

// 	MoveWindow(TabCtrl_1.hTab,0,0,(rc.right - rc.left)/2-4,rc.bottom - rc.top,FALSE);
// 	for( i=0;i<TabCtrl_1.tabPageCount;i++)
// 		TabCtrl_1.StretchTabPage(TabCtrl_1.hTab,i);

	MoveWindow(g_MainWebTabCtrl.hTab, 0, iHight, (rc.right - rc.left)-4, rc.bottom - rc.top - iHight, FALSE);
	for( i=0;i<g_MainWebTabCtrl.tabPageCount;i++)
		//gMainWebTabCtrl.CenterTabPage(gMainWebTabCtrl.hTab,i);
		g_MainWebTabCtrl.StretchTabPage(g_MainWebTabCtrl.hTab,i);

	// Refresh(hwnd);
	RedrawWindow(hwnd,NULL,NULL,RDW_ERASE|RDW_INVALIDATE|RDW_ALLCHILDREN|RDW_UPDATENOW);
}
// 
// void TabCtrl1_TabPages_OnSize(HWND hwnd, UINT state, int cx, int cy)
// {
// 	/////////////////////////////////////////////////////
// 	// Sizeing and positioning of Tab Page Widgets shall
// 	// be handled here so that they remain consistent with
// 	// the tab page.
// 	// Remember that the hwnd changes with each successive tab page
// 
// 	RECT  rc, chRc;
// 	int h, w;
// 	GetClientRect(hwnd, &rc);
// 
// 	if(hwnd==TabCtrl_1.hTabPages[0])
// 	{
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_1),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_1),rc.left+(rc.right-rc.left-w)/2,rc.top+(rc.bottom-rc.top)/4-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_2),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_2),rc.left+(rc.right-rc.left-w)/2,rc.top+(rc.bottom-rc.top)/2-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_3),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_3),rc.left+(rc.right-rc.left-w)/2,rc.top+(rc.bottom-rc.top)/4*3-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 	}
// 	else if(hwnd==TabCtrl_1.hTabPages[1])
// 	{
// 		MoveWindow(GetDlgItem(hwnd,FRA_TAB_PAGE_2),4,0,rc.right - rc.left-4,rc.bottom - rc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_4),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_4),rc.left+(rc.right-rc.left-w)/3*2,rc.top+(rc.bottom-rc.top)/4-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,LBL_1),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,LBL_1),rc.left+(rc.right-rc.left-w)/2,rc.top+(rc.bottom-rc.top)/2-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_5),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_5),rc.left+(rc.right-rc.left-w)/3,rc.top+(rc.bottom-rc.top)/4*3-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 	}
// 	else if(hwnd==TabCtrl_1.hTabPages[2])
// 	{
// 		MoveWindow(GetDlgItem(hwnd,FRA_TAB_PAGE_3),4,0,rc.right - rc.left-4,rc.bottom - rc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_6),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_6),rc.left+(rc.right-rc.left-w)/3,rc.top+(rc.bottom-rc.top)/4-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,LBL_2),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,LBL_2),rc.left+(rc.right-rc.left-w)/2,rc.top+(rc.bottom-rc.top)/2-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 
// 		GetWindowRect(GetDlgItem(hwnd,CMD_CLICK_7),&chRc);
// 		h=chRc.bottom-chRc.top;
// 		w=chRc.right-chRc.left;
// 		MoveWindow(GetDlgItem(hwnd,CMD_CLICK_7),rc.left+(rc.right-rc.left-w)/3*2,rc.top+(rc.bottom-rc.top)/4*3-h/2,chRc.right - chRc.left,chRc.bottom - chRc.top,FALSE);
// 	}
// 	else //hwnd==TabCtrl_1.hTabPages[3]
// 	{
// 		MoveWindow(GetDlgItem(hwnd,FRA_TAB_PAGE_4),4,0,rc.right - rc.left-4,rc.bottom - rc.top,FALSE);
// 	}
// }

void Main_TabCtrl_TabPages_OnSize(HWND hwnd, UINT state, int cx, int cy)
{
	/////////////////////////////////////////////////////
	// Sizeing and positioning of Tab Page Widgets shall
	// be handled here so that they remain consistent with
	// the tab page.
	// Remember that the hwnd changes with each successive tab page
	RECT	rect;
	GetClientRect(hwnd, &rect);

	ResizeBrowser( hwnd, rect.right, rect.bottom );

}


LRESULT  FormMain_OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{	
	RECT rc;
	HWND hWndEdit = NULL;
	static LPWSTR WebTaNames[]= {L"xx", 0};
	static LPWSTR WebDlgNames[]= {MAKEINTRESOURCE(TAB_CONTROL_2_PAGE_1),0};

	InitHandles (hwnd); // this line must be first!


// 
// 	New_TabControl( &TabCtrl_1, // address of TabControl struct
// 					GetDlgItem(hwnd, TAB_CONTROL_1), // handle to tab control
// 					tabnames, // text for each tab
// 					dlgnames, // dialog id's of each tab page dialog
// 					&FormMain_DlgProc, // address of main windows proc
// 					&TabCtrl1_TabPages_OnSize, // optional address of size function or NULL
// 					TRUE); // stretch tab page to fit tab ctrl
// 


	CHtmlNewTabControl( &g_MainWebTabCtrl,
					GetDlgItem(hwnd, MAIN_WEB_TAB_CONTROL),
					WebTaNames,
					WebDlgNames,
					&FormMain_DlgProc,
					&Main_TabCtrl_TabPages_OnSize,
					TRUE);

// 	Edit_SetText(GetDlgItem(gMainWebTabCtrl.hTabPages[0],LBL_3),
// 		"Use the Tab key to navagate between\r\ntab controls.\r\n\r\n"
// 		"Use the arrow keys or PageUp/Down keys\r\nto select a tab.\r\n\r\n"
// 		"Use the arrow keys to enter a tab page.\r\n\r\n"
// 		"Use the Tab key or arrow keys to navagate\r\ntab page child commands.\r\n\r\n"
// 		"Use Enter key to execute selected command.\r\n\r\n"
// 		"Try resizing the dialog while looking\r\nat different tabs.");

	//SetWindowTextW(GetDlgItem(hwnd, IDC_EDIT_URL), L"http://");
	
	//----Everything else must follow the above----//

	//Get the initial Width and height of the dialog
	//in order to fix the minimum size of dialog
	
	GetWindowRect(hwnd,&rc);
	g_MinSize.cx = (rc.right - rc.left);
	g_MinSize.cy = (rc.bottom - rc.top);

	hWndEdit=GetDlgItem( hwnd,IDC_EDIT_URL );

    wpFormMain_OrigEditProc = (WNDPROC) SetWindowLong(hWndEdit,GWL_WNDPROC,(LONG) FormMain_EditProc);

	return 0;
}




void FormMain_OnGetMinMaxInfo(HWND hwnd, LPMINMAXINFO lpMinMaxInfo)
{
    lpMinMaxInfo->ptMinTrackSize.x = g_MinSize.cx; //<-- Min. width (in pixels) of your window
    lpMinMaxInfo->ptMinTrackSize.y = g_MinSize.cy; //<-- Min. height (in pixels) of your window

/*  lpMinMaxInfo->ptMaxTrackSize.x = 640; //<-- Max. width (in pixels) of your window
    lpMinMaxInfo->ptMaxTrackSize.y = 480; //<-- Max. height (in pixels) of your window
    lpMinMaxInfo->ptMaxPosition.x = 0;    //<-- Left position (in pixels) of your window when maximized
    lpMinMaxInfo->ptMaxPosition.y = 0;    //<-- Top position (in pixels) of your window when maximized
 */
}
BOOL FormMain_OnCreateNewIEView(WPARAM wParam,LPARAM lParam)
{

	return CHtmlTabControl_OnCreateNewIEView(g_MainWebTabCtrl.hTab, wParam, lParam);
}

BOOL FormMain_OnNcLButtonDown(HWND hwnd, BOOL fDoubleClick, int x, int y, UINT codeHitTest)
{
	//SendMessage(g_frmMain, WM_NCLBUTTONDOWN, HTCAPTION, MAKELPARAM(x, y)); 
	//SendMessage(g_frmMain, WM_SYSCOMMAND,SC_MOVE + HTCAPTION ,0);
	return FALSE;
}
/****************************************************************************
 *                                                                          *
 * Function: MainDlgProc                                                    *
 *                                                                          *
 * Purpose : Process messages for the Main dialog.                          *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

LRESULT CALLBACK FormMain_DlgProc (HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch(msg)
	{
		HANDLE_MSG (hwndDlg, WM_CLOSE, FormMain_OnClose);
		HANDLE_MSG (hwndDlg, WM_COMMAND, FormMain_OnCommand);
		HANDLE_MSG (hwndDlg, WM_INITDIALOG, FormMain_OnInitDialog);
		HANDLE_MSG (hwndDlg, WM_SIZE, FormMain_OnSize);
		HANDLE_MSG (hwndDlg, WM_GETMINMAXINFO, FormMain_OnGetMinMaxInfo);
		HANDLE_MSG (hwndDlg, WM_NOTIFY, FormMain_OnNotify);
		HANDLE_MSG (hwndDlg, WM_NCLBUTTONDOWN, FormMain_OnNcLButtonDown);
	case WM_NEW_IEVIEW :
		return (LRESULT)FormMain_OnCreateNewIEView( wParam, lParam );
		break;

		//// TODO: Add dialog message crackers here...

	default: return 0;
	}
	return 0;
}

/****************************************************************************
 *                                                                          *
 * Function: WinMain                                                        *
 *                                                                          *
 * Purpose : Initialize the application.  Register a window class,          *
 *           create and display the main window and enter the               *
 *           message loop.                                                  *
 *                                                                          *
 * History : Date      Reason                                               *
 *           00/00/00  Created                                              *
 *                                                                          *
 ****************************************************************************/

int PASCAL WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpszCmdLine, int nCmdShow)
{
    WNDCLASSEX wcx;

    g_hInstance = hInstance;

    // Initialize common controls. Also needed for MANIFEST's.
    InitCommonControls();

    // Load Rich Edit control support.
    // TODO: uncomment one of the lines below, if you are using a Rich Edit control.
    // LoadLibrary(_T("riched32.dll"));  // Rich Edit v1.0
    // LoadLibrary(_T("riched20.dll"));  // Rich Edit v2.0, v3.0

    // Get system dialog information.
    wcx.cbSize = sizeof(wcx);
    if (!GetClassInfoEx(NULL, MAKEINTRESOURCE(32770), &wcx))
        return 0;
	if (OleInitialize(NULL) != S_OK)
        return 0;	

    // Add our own stuff.
    wcx.hInstance = hInstance;
    wcx.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDR_ICO_MAIN));
    wcx.lpszClassName = _T("TabCtrlDmoClass");
    if (!RegisterClassEx(&wcx))
        return 0;

    // The user interface is a modal dialog box.
    return DialogBox(hInstance, MAKEINTRESOURCE(FRM_MAIN), NULL, (DLGPROC)FormMain_DlgProc);
}


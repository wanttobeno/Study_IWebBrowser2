#pragma   once

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <malloc.h> 
#include "CHTMLTabCtrl.h"
#include "resource.h"
#include <OAIDL.H>
#include <mshtmhst.h>		// Defines of stuff like IDocHostUIHandler. This is an include file with Visual C 6 and above
#include <mshtml.h>			// Defines of stuff like IHTMLDocument2. This is an include file with Visual C 6 and above
#include <exdisp.h>			// Defines of stuff like IWebBrowser2. This is an include file with Visual C 6 and above


#include <crtdbg.h>			// for _ASSERT()

#include "misc.cpp"			// 杂项的工具类文件
#include "CWebevent.c"
#include "CWebControl.c"


/****************************************************************************/
// Internal Globals

#define CMD_VK_ESCAPE	101
#define CMD_VK_RETURN	102
static LPTABCTRL CHtml_m_lptc;
static BOOL CHtmlstopTabPageMessageLoop=FALSE;
_IDispatchEx CHtml_m_xEventSink;

// 当前激活的tab，判断是否可以关闭窗口
static HWND g_ActiveTabHwnd = 0;

// The VTable for our _IDispatchEx object.
//************************************
// Method:    Dispatch_GetTypeInfoCount
// FullName:  Dispatch_GetTypeInfoCount
// Access:    public 
// Returns:   IDispatchVtbl
// Qualifier:
//************************************
IDispatchVtbl	MyDispatch=
{
	IDispEventSimpleImpl_QueryInterface,
	IDispEventSimpleImpl_AddRef,
	IDispEventSimpleImpl_Release,
	IDispEventSimpleImpl_GetTypeInfoCount,
	IDispEventSimpleImpl_GetTypeInfo,
	IDispEventSimpleImpl_GetIDsOfNames,
	IDispEventSimpleImpl_Invoke
};
/****************************************************************************/
// Local function prototypes

static VOID CHtmlTabPageMessageLoop (HWND);
VOID CHtmlResetTabPageMessageLoop (HWND);

/****************************************************************************
*
*     FUNCTION: TabControl_GetClientRect
*
*     PURPOSE:  Return a Client rectangle for the tab control under
*               every possible configuration (tabs, buttons vert etc.)
*
*     PARAMS:   HWND hwnd - handle of a tab control
*               RECT *prc - pointer to a rectangle structure 
*
*     RETURNS:  VOID
*
*     COMMENTS: The rectangle structure is populated as follows:
*               prc.left = left, prc.top = top, prc.right = WIDTH, and prc.bottom = HEIGHT 
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlTabControl_GetClientRect(HWND hwnd,RECT* prc)
{
	RECT rtab_0;
	LONG lStyle = GetWindowLong(hwnd,GWL_STYLE); 

	// Calculate the tab control's display area
	GetWindowRect(hwnd, prc);
	ScreenToClient(GetParent(hwnd), (POINT*)&prc->left);
	ScreenToClient(hwnd, (POINT*)&prc->right);
	TabCtrl_GetItemRect(hwnd,0,&rtab_0); //The tab itself

	if((lStyle & TCS_BOTTOM) && (lStyle & TCS_VERTICAL)) //Tabs to Right
	{
		prc->top = prc->top + 6; //x coord
		prc->left = prc->left + 4; //y coord
		prc->bottom = prc->bottom - 12; // height
		prc->right = prc->right - (12 + rtab_0.right-rtab_0.left); // width
	}
	else if(lStyle & TCS_VERTICAL) //Tabs to Left
	{
		prc->top = prc->top + 6; //x coord
		prc->left = prc->left + (4 + rtab_0.right-rtab_0.left); //y coord
		prc->bottom = prc->bottom - 12; // height
		prc->right = prc->right - (12 + rtab_0.right-rtab_0.left); // width
	}
	else if(lStyle & TCS_BOTTOM) //Tabs on Bottom
	{
		prc->top = prc->top + 6; //x coord
		prc->left = prc->left + 4; //y coord
		prc->bottom = prc->bottom - (16 + rtab_0.bottom-rtab_0.top); // height
		prc->right = prc->right - 12; // width
	}
	else //Tabs on top
	{
		prc->top = prc->top + (6 + rtab_0.bottom-rtab_0.top); //x coord
		prc->left = prc->left + 4; //y coord
		prc->bottom = prc->bottom - (16 + rtab_0.bottom-rtab_0.top); // height
		prc->right = prc->right - 12; // width
	}
}

/****************************************************************************
*
*     FUNCTION: CHtmlCenterTabPage
*
*     PURPOSE:  Center the tab page in the tab control's display area.
*
*     PARAMS:   HWND hTab - handle tab control
*               INT iPage - index of tab page
*
*     RETURNS:  BOOL - TRUE if successful
*
*
* History:
*                June '06 - Created
*
\****************************************************************************/

BOOL CHtmlCenterTabPage (HWND hTab, INT iPage)
{
	RECT rect, rclient;

	// Update CHtml_m_lptc pointer
	CHtml_m_lptc = (LPTABCTRL) GetWindowLong(hTab,GWL_USERDATA);

 	CHtmlTabControl_GetClientRect(hTab, &rect); // left, top, width, height

	// Get the tab page size
	GetClientRect(CHtml_m_lptc->hTabPages[iPage], &rclient);
	rclient.right=rclient.right-rclient.left;// width
	rclient.bottom=rclient.bottom-rclient.top;// height
	rclient.left= rect.left;
	rclient.top= rect.top;

	// Center the tab page, or cut it off at the edge of the tab control(bad)
	if(rclient.right<rect.right)
		rclient.left += (rect.right-rclient.right)/2;

	if(rclient.bottom<rect.bottom)
		rclient.top += (rect.bottom-rclient.bottom)/2;

	// Move the child and put it on top
	return SetWindowPos(CHtml_m_lptc->hTabPages[iPage], HWND_TOP,
			rclient.left, rclient.top, rclient.right, rclient.bottom,
			0);
}

/****************************************************************************
*
*     FUNCTION: CHtmlStretchTabPage
*
*     PURPOSE:  Stretch the tab page to fit the tab control's display area.
*
*     PARAMS:   HWND hTab - handle tab control
*               INT iPage - index of tab page
*
*     RETURNS:  BOOL - TRUE if successful
*
*
* History:
*                June '06 - Created
*
\****************************************************************************/

BOOL CHtmlStretchTabPage (HWND hTab, INT iPage)
{
	RECT rect;

	// Update CHtml_m_lptc pointer
	CHtml_m_lptc = (LPTABCTRL) GetWindowLong(hTab,GWL_USERDATA);

 	CHtmlTabControl_GetClientRect(hTab, &rect); // left, top, width, height

	// Move the child and put it on top
	return SetWindowPos(CHtml_m_lptc->hTabPages[iPage], HWND_TOP,
			rect.left, rect.top, rect.right, rect.bottom,
			0);
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabCtrl_OnKeyDown
*
*     PURPOSE:  Handle key presses in the tab control (but not the tab pages).
*
*     PARAMS:   LPARAM lParam - lParam furnished by parent proc's WM_NOTIFY handler
*
*     RETURNS:  VOID
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlTabCtrl_OnKeyDown(LPARAM lParam)
{
	TC_KEYDOWN *tk=(TC_KEYDOWN *)lParam;
	int itemCount=TabCtrl_GetItemCount(tk->hdr.hwndFrom);
	int currentSel=TabCtrl_GetCurSel(tk->hdr.hwndFrom);
	BOOL verticalTabs ;

	if(itemCount <= 1) return; // Ignore if only one TabPage

	verticalTabs = GetWindowLong(CHtml_m_lptc->hTab, GWL_STYLE) & TCS_VERTICAL;
	
	if(verticalTabs)
	{
		switch (tk->wVKey)
		{
			case VK_PRIOR: //select the previous page
			{
				if(0==currentSel) return;
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel-1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom,currentSel-1);
			}
			break;
			case VK_UP: //select the previous page
			{
				if(0==currentSel) return;
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel-1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom, currentSel);
			}
			break;
			case VK_NEXT:  //select the next page
			{
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel+1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom, currentSel+1);
			}
			break;
			case VK_DOWN: //select the next page
			{
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel+1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom,currentSel);
			}
			break;
			case VK_LEFT: //navagate within selected child tab page
			{
				SetFocus(CHtml_m_lptc->hTabPages[currentSel]); // focus to child tab page
				CHtmlTabPageMessageLoop (CHtml_m_lptc->hTabPages[currentSel]); //start message loop
			}
			break;
			case VK_RIGHT: //navagate within selected child tab page
			{
				SetFocus(CHtml_m_lptc->hTabPages[currentSel]);
				CHtmlTabPageMessageLoop (CHtml_m_lptc->hTabPages[currentSel]);
			}
			break; 
			default: return;
		}
	} // if(verticalTabs)

	else // horizontal Tabs
    {
		switch (tk->wVKey)
		{
			case VK_NEXT: //select the previous page
			{
				if(0==currentSel) return;
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel-1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom, currentSel-1);
			}
			break;
			case VK_LEFT: //select the previous page
			{
				if(0==currentSel) return;
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel-1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom, currentSel);
			}
			break;
			case VK_PRIOR: //select the next page
			{
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel+1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom,currentSel+1);
			}
			break;
			case VK_RIGHT: //select the next page
			{
				TabCtrl_SetCurSel(tk->hdr.hwndFrom, currentSel+1);
				TabCtrl_SetCurFocus(tk->hdr.hwndFrom,currentSel);
			}
			break;
			case VK_UP: //navagate within selected child tab page
			{
				SetFocus(CHtml_m_lptc->hTabPages[currentSel]);
				CHtmlTabPageMessageLoop (CHtml_m_lptc->hTabPages[currentSel]);
			}
			break;
			case VK_DOWN: //navagate within selected child tab page
			{
				SetFocus(CHtml_m_lptc->hTabPages[currentSel]);
				CHtmlTabPageMessageLoop (CHtml_m_lptc->hTabPages[currentSel]);
			}
			break; 
			default: return;
		}
	} //else // horizontal Tabs
}

/****************************************************************************
*
*     FUNCTION: CHtmlCreateAccTable
*
*     PURPOSE:  Creates a table of keyboard accelerators
*
*     PARAMS:   VOID

*     RETURNS:  HACCEL - The handle to the created table
*
* History:
*                June '06 - Created
*
\****************************************************************************/

HACCEL CHtmlCreateAccTable (VOID)
{
	static  ACCEL  aAccel[2];
	static  HACCEL  hAccel;

	aAccel[0].fVirt=FVIRTKEY;
	aAccel[0].key=VK_ESCAPE;
	aAccel[0].cmd=CMD_VK_ESCAPE;

	aAccel[1].fVirt=FVIRTKEY;
	aAccel[1].key=VK_RETURN;
	aAccel[1].cmd=CMD_VK_RETURN;

	hAccel=CreateAcceleratorTable(aAccel,1);
	return hAccel;
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabPage_OnCommand
*
*     PURPOSE:  Handle the Dialog virtual keys while
*               forwarding the rest of the commands to ParentProc.
*
*     PARAMS:   HWND hwnd - handle to the tab page (dialog)
*               INT id - Control id
*               HWND hwndCtl - Controls handle
*               UINT codeNotify - Notification code 
*
*     RETURNS:  VOID
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlTabPage_OnCommand(HWND hwnd, INT id, HWND hwndCtl, UINT codeNotify)
{
	if(CMD_VK_ESCAPE==id)
	{
		CHtmlstopTabPageMessageLoop=TRUE; // cause message loop to return
		SetFocus(CHtml_m_lptc->hTab); // focus to tab control
		return;
	}
	else if(CMD_VK_RETURN==id)
	{
		//
		// TODO: We may want to handle this
		//  ex: If we are editing a tree view on a tab,
		//  this will make the Enter key work
		//  TreeView_EndEditLabelNow(hTree,FALSE);
		//
	}

	//Forward all other commands
	FORWARD_WM_COMMAND (hwnd,id,hwndCtl,codeNotify,CHtml_m_lptc->ParentProc);

	// Mouse clicks on a control should engage the Message Loop

	// If this WM_COMMAND message is a notification to parent window
	// ie: EN_SETFOCUS being sent when an edit control is initialized
	// do not engage the Message Loop or send any messages.
	if(codeNotify!=0) return;

	CHtmlResetTabPageMessageLoop (hwnd);

	// Toggling WM_NEXTDLGCTL ensures that default focus moves to selected
	// Control (This must follow call to ResetTabPageMessageLoop()) 
	SendMessage(hwnd, WM_NEXTDLGCTL, (WPARAM)0, FALSE);
	SendMessage(hwnd, WM_NEXTDLGCTL, (WPARAM)1, FALSE);
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabPage_OnLButtonDown
*
*     PURPOSE:  Handle WM_LBUTTONDOWN (tab page dialog)
*
*     PARAMS:   HWND hwnd - handle to the tab page (dialog)
*               BOOL fDoubleClick - Was this a double click?
*               INT x - Coordinate
*               INT y - Coordinate
*               UINT keyFlags - Flags 
*
*     RETURNS:  VOID
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlTabPage_OnLButtonDown(HWND hwnd, BOOL fDoubleClick, INT x, INT y, UINT keyFlags)
{
	// If Mouse click in tab page but not on control
	CHtmlResetTabPageMessageLoop (hwnd);
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabPage_OnSize
*
*     PURPOSE:  Handle WM_SIZE (tab page dialog)
*
*     PARAMS:   HWND hwnd - handle to the tab page (dialog)
*               UINT state - resizing flag
*               INT cx - Width
*               INT cy - Height
*
*     RETURNS:  VOID
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlTabPage_OnSize(HWND hwnd, UINT state, INT cx, INT cy)
{
	// ResizeBrowser(hwnd, LOWORD(lParam), HIWORD(lParam));
	// Dummy function when no external on size function is
	// desired.
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabPage_OnInitDialog
*
*     PURPOSE:  Handle WM_INITDIALOG (tab page dialog)
*
*     PARAMS:   HWND hwnd - handle to the tab page (dialog)
*               HWND hwndFocus - handle of control to receive focus
*               LPARAM lParam - initialization parameter
*
*     RETURNS:  BOOL - TRUE if handled
*
* History:
*                June '06 - Created
*
\****************************************************************************/

BOOL CHtmlTabPage_OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
	IID	iid;

	// Embed the browser object into our host window.

	if (EmbedBrowserObject(hwnd)) return(-1);
				
	COleControlSite__GetEventIID( hwnd, &iid);
				
	CHtml_m_xEventSink.dispatchObj = (IDispatch *)&MyDispatch;
	// memcpy(&CHtml_m_xEventSink.dispatchObj, &MyDispatch, sizeof(IDispatch) );
	COleControlSite__ConnectSink( hwnd, &iid, (IUnknown *)&CHtml_m_xEventSink );
				
	return DefWindowProc(hwnd, WM_INITDIALOG, (WPARAM)hwndFocus, lParam);
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabPage_DlgProc
*
*     PURPOSE:  Window Procedure for the tab pages (dialogs).
*
*     PARAMS:   HWND   hwnd		- This window
*               UINT   msg		- Which message?
*               WPARAM wParam	- message parameter
*               LPARAM lParam	- message parameter
*
*     RETURNS:  BOOL CALLBACK - depends on message
*
* History:
*                June '06 - Created
*
\****************************************************************************/

BOOL CALLBACK CHtmlTabPage_DlgProc (HWND hwndDlg, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch(msg)
	{
		HANDLE_MSG (hwndDlg, WM_INITDIALOG, CHtmlTabPage_OnInitDialog);
		HANDLE_MSG (hwndDlg, WM_SIZE, CHtml_m_lptc->TabPage_OnSize);
		HANDLE_MSG (hwndDlg, WM_COMMAND, CHtmlTabPage_OnCommand);
		HANDLE_MSG (hwndDlg, WM_LBUTTONDOWN, CHtmlTabPage_OnLButtonDown);
		//// TODO: Add TabPage dialog message crackers here...
		default: return FALSE; //CHtml_m_lptc->ParentProc (hwndDlg, msg, wParam, lParam);
	}
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabPageMessageLoop
*
*     PURPOSE:  Monitor and respond to user keyboard input and system messages.
*
*     PARAMS:   HWND hwnd - handle to the currently visible tab page
*
*     RETURNS:  VOID
*
*     COMMENTS: Send PostQuitMessage(0); from any cancel or exit event.	
*               Failure to do so will leave the process running even after
*               application exit.
*
* History:
*                June '06 - Created
*
\****************************************************************************/

static VOID CHtmlTabPageMessageLoop (HWND hwnd)
{
	MSG	msg;
	int	status;
	BOOL handled = FALSE;
	HACCEL hAccTable = CHtmlCreateAccTable();

	while((status = GetMessage(&msg, NULL, 0, 0 )) != 0 && !CHtmlstopTabPageMessageLoop)
	{ 
	    if (status == -1) // Exception
	    {
	        return;
	    }
		else
	    {
			// Dialogs do not expose a WM_KEYDOWN message so we will seperate
			// the desired keyboard events here
			handled = TranslateAccelerator(hwnd,hAccTable,&msg);

			// Perform default dialog message processing using IsDialogM. . .
			if(!handled) handled=IsDialogMessage(hwnd,&msg);

			// Non dialog message handled in the standard way.
	        if(!handled)
	        {
	            TranslateMessage(&msg);
	            DispatchMessage(&msg);
	        }
	    }
	}
	if(CHtmlstopTabPageMessageLoop) //Reset: do not PostQuitMessage(0)
	{
		DestroyAcceleratorTable(hAccTable);
		CHtmlstopTabPageMessageLoop = FALSE;
		return;
	}

	// Default: Re-post the Quit message
	DestroyAcceleratorTable(hAccTable);
	PostQuitMessage(0);
	return;
}

/****************************************************************************
*
*     FUNCTION: ResetTabPageMessageLoop
*
*     PURPOSE:  Toggles message loop so that it is active only in the tab page.
*
*     PARAMS:   HWND hwnd - handle to the currently visible tab page
*
*     RETURNS:  VOID
*
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlResetTabPageMessageLoop (HWND hwnd)
{
	//Toggle kill sw
	CHtmlstopTabPageMessageLoop=TRUE;
	CHtmlstopTabPageMessageLoop=FALSE;

	SetFocus(hwnd);
	CHtmlTabPageMessageLoop(hwnd);
}

/****************************************************************************
*
*     FUNCTION: CHtmlTabCtrl_OnSelChanged
*
*     PURPOSE:  A tab has been pressed, Handle TCN_SELCHANGE notification.
*
*     PARAMS:   VOID
*
*     RETURNS:  BOOL - TRUE if handled
*
* History:
*                June '06 - Created
*
\****************************************************************************/

BOOL CHtmlTabCtrl_OnSelChanged(VOID)
{
	int iSel = TabCtrl_GetCurSel(CHtml_m_lptc->hTab);

	//Hide the current child dialog box, if any.
	ShowWindow(CHtml_m_lptc->hVisiblePage,FALSE);

	//Show the new child dialog box.
	ShowWindow(CHtml_m_lptc->hTabPages[iSel],TRUE);

	// Save the current child
	CHtml_m_lptc->hVisiblePage = CHtml_m_lptc->hTabPages[iSel];

	return TRUE;
}

/****************************************************************************
*
*     FUNCTION: CHtmlNotify
*
*     PURPOSE:  Handle WM_NOTIFY messages.
*
*     PARAMS:   VOID
*
*     RETURNS:  BOOL - TRUE if handled
*
*     COMMENTS: This function is made external to module via function pointer
*
* History:
*                June '07 - Created
*
\****************************************************************************/
BOOL CHtmlNotify (LPNMHDR pnm)
{
	// Update CHtml_m_lptc pointer

	CHtml_m_lptc = (LPTABCTRL) GetWindowLong(pnm->hwndFrom,GWL_USERDATA);

	switch (pnm->code)
	{
		case TCN_KEYDOWN:
			CHtmlTabCtrl_OnKeyDown((LPARAM)pnm);
			break;
			// fall through to call CHtmlTabCtrl_OnSelChanged() on each keydown
		//case NM_DBLCLK:
		//	return CHtmlTabControl_Close( );
		case NM_CLICK:	
			{
				if (g_ActiveTabHwnd == CHtml_m_lptc->hVisiblePage)
				{
					CHtmlTabControl_Close( );
					g_ActiveTabHwnd = CHtml_m_lptc->hVisiblePage;
				}
				else
				{
							g_ActiveTabHwnd = CHtml_m_lptc->hVisiblePage;
				}
				return TRUE;
			}
		 break;
		case TCN_SELCHANGE:
			// 切换后,比NM_CLICK先到
			g_ActiveTabHwnd = CHtml_m_lptc->hVisiblePage;
			CHtmlTabCtrl_OnSelChanged();
			return TRUE;
		//case TCN_SELCHANGING:
		//	return TRUE;
	}
	return FALSE;
}


BOOL CHtmlTabControl_DoAction( HWND hwnd, int iAction)
{
	IWebBrowser2	*webBrowser2;
	IOleObject		*browserObject;
	
	// Retrieve the browser object's pointer we stored in our window's GWL_USERDATA when
	// we initially attached the browser object to this window.
	browserObject = *((IOleObject **)GetWindowLong(CHtml_m_lptc->hVisiblePage, GWL_USERDATA));
	
	// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
	// object, so we can call some of the functions in the former's table.
	if (!browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**)&webBrowser2))
	{
		// ignore js error
		webBrowser2->lpVtbl->put_Silent(webBrowser2,VARIANT_TRUE);

		// Ok, now the pointer to our IWebBrowser2 object is in 'webBrowser2', and so its VTable is
		// webBrowser2->lpVtbl.
		
		// Call the desired function
		switch (iAction)
		{
		case IDC_BUTTON_GOBACK:
			{
				// Call the IWebBrowser2 object's GoBack function.
				webBrowser2->lpVtbl->GoBack(webBrowser2);
				break;
			}
			
		case IDC_BUTTON_GOFORWARD:
			{
				// Call the IWebBrowser2 object's GoForward function.
				webBrowser2->lpVtbl->GoForward(webBrowser2);
				break;
			}
			
			// 			case WEBPAGE_GOHOME:
			// 			{
			// 				// Call the IWebBrowser2 object's GoHome function.
			// 				webBrowser2->lpVtbl->GoHome(webBrowser2);
			// 				break;
			// 			}
			// 
			// 			case WEBPAGE_SEARCH:
			// 			{
			// 				// Call the IWebBrowser2 object's GoSearch function.
			// 				webBrowser2->lpVtbl->GoSearch(webBrowser2);
			// 				break;
			// 			}
			
		case IDC_BUTTON_FLUSH:
			{
				// Call the IWebBrowser2 object's Refresh function.
				webBrowser2->lpVtbl->Refresh(webBrowser2);
			}
			
		case IDC_BUTTON_STOP:
			{
				// Call the IWebBrowser2 object's Stop function.
				webBrowser2->lpVtbl->Stop(webBrowser2);
			}
		}
		
		// Release the IWebBrowser2 object.
		webBrowser2->lpVtbl->Release(webBrowser2);
	}
	return TRUE;
}


/****************************************************************************
*
*     FUNCTION: CHtmlTabControl_Destroy
*
*     PURPOSE:  Destroy the tab page dialogs and
*               free the list of pointers to the dialogs.
*
*     PARAMS:   LPTABCTRL tc - Pointer to a TABCTRL structure
*
*     RETURNS:  VOID
*
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlTabControl_Destroy(LPTABCTRL tc)
{
	int i=0;
	for ( i=0;i<tc->tabPageCount;i++)
		DestroyWindow(tc->hTabPages[i]);

	free (tc->hTabPages);
}

// WM_NEW_IEVIEW 的消息响应
BOOL CHtmlTabControl_OnCreateNewIEView(HWND hWnd, WPARAM wParam,LPARAM lParam)
{
	int		iErr	= 0;
	HWND	*hTab	= (HWND*)wParam;
	TCITEMW	tie;
//	NMHDR   nmhdr;   
	int i			= 0;
	int iSel = TabCtrl_GetCurSel(CHtml_m_lptc->hTab);

	CHtml_m_lptc = (LPTABCTRL) GetWindowLong(hWnd,GWL_USERDATA);

	i = CHtml_m_lptc->tabPageCount++;
	// Add a tab for each name in tabnames (list ends with 0)
	tie.mask = TCIF_TEXT | TCIF_IMAGE;
	tie.iImage = -1;
	
	tie.pszText = CHtml_m_lptc->tabNames[0];
	TabCtrl_InsertItem(CHtml_m_lptc->hTab, i, &tie);

	CHtml_m_lptc->hTabPages[i] = CreateDialog(
		GetModuleHandle(NULL),
		MAKEINTRESOURCE(TAB_CONTROL_2_PAGE_1),
		GetParent(CHtml_m_lptc->hTab),
		(DLGPROC)CHtmlTabPage_DlgProc);
	
	*hTab = CHtml_m_lptc->hTabPages[i];	
	// 	nmhdr.code = TCN_SELCHANGING;   
	// 	nmhdr.hwndFrom = gMainWebTabCtrl.hTab;   
	// 	nmhdr.idFrom = MAIN_WEB_TAB_CONTROL;   
	// 	SendMessage( 
	// 		gMainWebTabCtrl.hTab, 
	// 		WM_NOTIFY, 
	// 		MAKELONG(TCN_SELCHANGING,0), 
	// 		(LPARAM)(&nmhdr)
	// 		);   
	TabCtrl_SetCurSel( CHtml_m_lptc->hTab, i );
	
	if( CHtml_m_lptc->blStretchTabs )
	{
		CHtml_m_lptc->StretchTabPage( CHtml_m_lptc->hTab, i );
	}
	else
	{
		CHtml_m_lptc->CenterTabPage( CHtml_m_lptc->hTab, i );
	}


	CHtmlTabCtrl_OnSelChanged();

	g_ActiveTabHwnd = CHtml_m_lptc->hTabPages[i];
	return TRUE;
}

BOOL CHtmlTabControl_Close( )
{
	int iSel = TabCtrl_GetCurSel( CHtml_m_lptc->hTab );

	int iNextSel = 0;
	int i = 0;

	if(CHtml_m_lptc->hVisiblePage != CHtml_m_lptc->hTabPages[iSel] && (CHtml_m_lptc->tabPageCount != 1))
		return FALSE;

	iNextSel = iSel ? iSel-1 : 1;

	if ( CHtml_m_lptc->tabPageCount ) // 如果没有页了
	{
		CHtml_m_lptc->tabPageCount--;
	}else{
		return FALSE;
	}

	if ( CHtml_m_lptc->tabPageCount ) 
	{
		UnEmbedBrowserObject( CHtml_m_lptc->hVisiblePage );
		EndDialog( CHtml_m_lptc->hVisiblePage, 0 );
		TabCtrl_SetCurSel( CHtml_m_lptc->hTab, iNextSel );
		TabCtrl_DeleteItem( CHtml_m_lptc->hTab, iSel);

		if( CHtml_m_lptc->blStretchTabs )
		{
			CHtml_m_lptc->StretchTabPage( CHtml_m_lptc->hTab, iSel ); // 调整PAGE的大小
		}
		else
		{
			CHtml_m_lptc->CenterTabPage( CHtml_m_lptc->hTab, iSel );
		}
		
		ShowWindow(CHtml_m_lptc->hTabPages[iNextSel],TRUE);

		CHtml_m_lptc->hVisiblePage = CHtml_m_lptc->hTabPages[iNextSel];
		for ( i = iSel; i<CHtml_m_lptc->tabPageCount; i++)
		{
			CHtml_m_lptc->hTabPages[i] = CHtml_m_lptc->hTabPages[i+1];
		}
		CHtml_m_lptc->hTabPages[i] = 0;
		
	}else{ // 如果是最后一页了
		CHtml_m_lptc->tabPageCount++;
		DisplayHTMLPage(CHtml_m_lptc->hVisiblePage, L"about:blank");
	}


	return TRUE;
}
/****************************************************************************
*
*     FUNCTION: CHtmlNewTabControl
*
*     PURPOSE:  Initialize a TABCTRL struct, create tab page dialogs
*               and start the message loop.
*
*     PARAMS:   LPTABCTRL tc - Poinnnter to a TABCTRL structure
*               HWND hTab - Handle to the parent tab control
*               LPSTR *tabNames - Pointer to array of tab names
*               LPSTR *dlgNames - Pointer to array of MAKEINTRESOURCE() dialog ids
*               BOOL CALLBACK(*ParentProc)(HWND, UINT, WPARAM, LPARAM) - function pointer
*               VOID (*OnSize)(HWND, UINT, int, int) - Optional function pointer
*               BOOL fStretch) - stretch dialog flag
*
*     RETURNS:  VOID
*
*
* History:
*                June '06 - Created
*
\****************************************************************************/

VOID CHtmlNewTabControl(LPTABCTRL lptc,
					HWND hTab,
					LPWSTR *tabNames,
					LPWSTR *dlgNames,
					BOOL (CALLBACK* ParentProc)(HWND, UINT, WPARAM, LPARAM),
					VOID (*OnSize)(HWND, UINT, int, int),
					BOOL fStretch)
{
	
	LPWSTR* ptr = NULL;
	int i = 0;
	int iErr = 0;
	HWND hWnd = NULL;
	TCITEMW		tcHtmlTabControlTie;
	CHtml_m_lptc=lptc;

	//Link struct CHtml_m_lptc pointer to Main Tab ConCtrl
	SetWindowLong(hTab,GWL_USERDATA,(long)CHtml_m_lptc);

	CHtml_m_lptc->hTab=hTab;
	CHtml_m_lptc->tabNames=tabNames;
	CHtml_m_lptc->dlgNames=dlgNames;
	CHtml_m_lptc->blStretchTabs=fStretch;

	// Point to external functions
	CHtml_m_lptc->ParentProc=ParentProc;
	if(NULL!=OnSize) //external size function
		CHtml_m_lptc->TabPage_OnSize=OnSize;
	else //internal dummy size function
		CHtml_m_lptc->TabPage_OnSize=CHtmlTabPage_OnSize;

	// Point to internal public functions
	CHtml_m_lptc->Notify=&CHtmlNotify;
	CHtml_m_lptc->StretchTabPage=&CHtmlStretchTabPage;
	CHtml_m_lptc->CenterTabPage=&CHtmlCenterTabPage;

	// Determine number of tab pages to insert based on DlgNames
	CHtml_m_lptc->tabPageCount = 0;
	ptr=CHtml_m_lptc->dlgNames;
	while(*ptr++) 
		CHtml_m_lptc->tabPageCount++;

	//create array based on number of pages
	CHtml_m_lptc->hTabPages = (HWND*)malloc( 50 * sizeof(HWND*));// 改成静态的了。50个还不够？注意我没释放！
	memset(CHtml_m_lptc->hTabPages, 0, 50 * sizeof(HWND*) );
	// Add a tab for each name in tabnames (list ends with 0)
	tcHtmlTabControlTie.mask = TCIF_TEXT | TCIF_IMAGE;
	tcHtmlTabControlTie.iImage = -1;
	for ( i = 0; i< CHtml_m_lptc->tabPageCount; i++)
	{
		tcHtmlTabControlTie.pszText = CHtml_m_lptc->tabNames[0];
		TabCtrl_InsertItem(CHtml_m_lptc->hTab, i, &tcHtmlTabControlTie);

		// Add page to each tab
		CHtml_m_lptc->hTabPages[i] = CreateDialog(GetModuleHandle(NULL),
										  CHtml_m_lptc->dlgNames[i],
										  GetParent(CHtml_m_lptc->hTab),
										  (DLGPROC)CHtmlTabPage_DlgProc);

		hWnd = CHtml_m_lptc->hTabPages[i];

		// 此句加在这虽然重复了几次，不过这是为了响应某个事件
		CHtml_m_lptc->hVisiblePage = CHtml_m_lptc->hTabPages[0];

		DisplayHTMLPage(hWnd, L"http://news.baidu.com");

		// Set initial tab page position
		if(CHtml_m_lptc->blStretchTabs)
			CHtml_m_lptc->StretchTabPage(CHtml_m_lptc->hTab, i);
		else
			CHtml_m_lptc->CenterTabPage(CHtml_m_lptc->hTab, i);
	}
	
	g_ActiveTabHwnd = CHtml_m_lptc->hTabPages[0];
	
	// Show first tab
	ShowWindow(CHtml_m_lptc->hTabPages[0],SW_SHOW);
	UpdateWindow(hWnd);
	// Save the current child

}


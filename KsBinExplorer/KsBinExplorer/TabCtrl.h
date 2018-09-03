/****************************************************************************\
*			 
*     FILE:     TABCTRL.H
*
*     PURPOSE:  Code for managing tab control property page
*               dialogs header file
*
*     COMMENTS: This source is distributed in the hope that it will be useful,
*               but WITHOUT ANY WARRANTY; without even the implied warranty of
*               MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
*
*     Copyright 2006 David MacDermot.
*
*
* History:
*                JUNE '06 - Created
*                JULY '07 - line 35 & 54 changed BOOL (*ParentProc) etc... to
*                                                BOOL (CALLBACK *ParentProc)
*                           line 65 fixed a bug in TabControl_SelectTab() macro
*
\****************************************************************************/

/****************************************************************************/
// Structs

typedef struct tagTabControl
{
	HWND hTab;
	HWND hVisiblePage;
	HWND* hTabPages;
	LPWSTR *tabNames;
	LPWSTR *dlgNames;
	int tabPageCount;
	BOOL blStretchTabs;

	// Function pointer to Parent Dialog Proc
	BOOL (CALLBACK *ParentProc)(HWND, UINT, WPARAM, LPARAM);

	// Function pointer to Tab Page Size Proc
	void (*TabPage_OnSize)(HWND hwnd, UINT state, int cx, int cy);

	// Pointers to shared functions
	BOOL (*Notify) (LPNMHDR);
	BOOL (*StretchTabPage) (HWND, INT);
	BOOL (*CenterTabPage) (HWND, INT);
 
}TABCTRL, *LPTABCTRL;

/****************************************************************************/
// Exported function prototypes

void New_TabControl(LPTABCTRL,
					HWND,
					LPWSTR*,
					LPWSTR*,
					BOOL (CALLBACK *ParentProc)(HWND, UINT, WPARAM, LPARAM),
					VOID (*TabPage_OnSize)(HWND, UINT, int, int),
					BOOL fStretch);

void TabControl_Destroy(LPTABCTRL);

/****************************************************************************/
// Macroes

#define TabCtrl_SelectTab(hTab,iSel) { \
	TabCtrl_SetCurSel(hTab,iSel); \
	NMHDR nmh = { hTab, GetDlgCtrlID(hTab), TCN_SELCHANGE }; \
	SendMessage(nmh.hwndFrom,WM_NOTIFY,(WPARAM)nmh.idFrom,(LPARAM)&nmh); }

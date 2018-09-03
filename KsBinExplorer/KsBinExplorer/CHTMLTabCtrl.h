

/****************************************************************************/
// Exported function prototypes

void CHtmlNewTabControl(LPTABCTRL a,
					HWND b, 
					LPWSTR* c,
					LPWSTR* d,
					BOOL (CALLBACK *ParentProc)(HWND, UINT, WPARAM, LPARAM),
					VOID (*TabPage_OnSize)(HWND, UINT, int, int),
					BOOL fStretch);

void CHtmlTabControlDestroy(LPTABCTRL);
BOOL CHtmlTabControl_Close();
/****************************************************************************/
// Macroes

#define TabCtrl_SelectTab(hTab,iSel) { \
	TabCtrl_SetCurSel(hTab,iSel); \
	NMHDR nmh = { hTab, GetDlgCtrlID(hTab), TCN_SELCHANGE }; \
	SendMessage(nmh.hwndFrom,WM_NOTIFY,(WPARAM)nmh.idFrom,(LPARAM)&nmh); }

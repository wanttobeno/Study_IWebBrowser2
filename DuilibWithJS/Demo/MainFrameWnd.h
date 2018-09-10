#pragma once
#include "WebBrowserExUI.h"


class CMainFrameWnd :
	public WindowImplBase
{
public:
	CMainFrameWnd();
	~CMainFrameWnd();

	virtual CDuiString GetSkinFile()	{ return _T("MainWnd.xml"); }
	virtual LPCTSTR GetWindowClassName(void) const	{ return _T("MainFrameWnd"); }

	virtual CControlUI* CreateControl(LPCTSTR pstrClass);
	virtual void InitWindow();
	virtual void OnClick(TNotifyUI& msg);
	virtual void OnFinalMessage(HWND hWnd);

	LRESULT OnDestroy(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	CWebBrowserExUI *m_pBrowser;
	CButtonUI *m_pBtn;
};


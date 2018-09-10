#include "stdafx.h"
#include "MainFrameWnd.h"
#include "MyWebBrowserEvenrHandler.h"
#include <atlcomcli.h>

CMainFrameWnd *g_pMainWnd;


CMainFrameWnd::CMainFrameWnd()
{
	g_pMainWnd = this;
}


CMainFrameWnd::~CMainFrameWnd()
{
}


CControlUI* CMainFrameWnd::CreateControl(LPCTSTR pstrClass)
{
	if (_tcsicmp(pstrClass, _T("WebBrowserEx")) == 0)
	{
		return new CWebBrowserExUI;
	}

	return NULL;
}


void CMainFrameWnd::InitWindow()
{
	CenterWindow();
	m_pBrowser = (CWebBrowserExUI*)m_pm.FindControl(_T("web")); ASSERT(m_pBrowser);
	m_pBtn = (CButtonUI*)m_pm.FindControl(_T("CallJsBtn")); ASSERT(m_pBtn);

	static CMyWebBrowserEvenrHandler Handler;
	m_pBrowser->SetWebBrowserEventHandler(&Handler);
	CDuiString strHtmlPath = CPaintManagerUI::GetInstancePath();
	strHtmlPath.Append(_T("test.html"));
	strHtmlPath.Replace(_T("\\"), _T("/"));
	CDuiString strUrl;
	strUrl.Format(_T("file:///%s"), strHtmlPath.GetData());
	m_pBrowser->NavigateUrl(strUrl.GetData());
}


void CMainFrameWnd::OnClick(TNotifyUI& msg)
{
	CDuiString sCtrlName = msg.pSender->GetName();

	if (sCtrlName.CompareNoCase(_T("CallJsBtn")) == 0)
	{
		// C++调用js方法，示例
		// 注意参数顺序，反向
		VARIANT arg[2] = { CComVariant(7),CComVariant(3)};//JsFunSub(3,7)
		VARIANT varRet;
		m_pBrowser->InvokeMethod(m_pBrowser->GetHtmlWindow(),_T("JsFunSub"),&varRet,arg,2);
		int nResult =  varRet.intVal;//-4
		return;
	}


	WindowImplBase::OnClick(msg);
}


void CMainFrameWnd::OnFinalMessage(HWND hWnd)
{
	__super::OnFinalMessage(hWnd);

	delete this;
}


LRESULT CMainFrameWnd::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	// 退出程序
	PostQuitMessage(0);

	bHandled = FALSE;
	return 0;
}
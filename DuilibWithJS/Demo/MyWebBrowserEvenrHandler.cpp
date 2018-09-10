#include "stdafx.h"
#include "MyWebBrowserEvenrHandler.h"
#include "MainFrameWnd.h"
#include <atlcomcli.h>
extern CMainFrameWnd *g_pMainWnd;


CMyWebBrowserEvenrHandler::CMyWebBrowserEvenrHandler()
{
}


CMyWebBrowserEvenrHandler::~CMyWebBrowserEvenrHandler()
{
}


void CMyWebBrowserEvenrHandler::BeforeNavigate2(CWebBrowserUI* pWeb, IDispatch* pDisp, VARIANT*& url, VARIANT*& Flags, VARIANT*& TargetFrameName, VARIANT*& PostData, VARIANT*& Headers, VARIANT_BOOL*& Cancel)
{
	CDuiString strUrl(url->bstrVal);
	//MyOutputDebugStringW(L"BeforeNavigate2:%s\r\n", strUrl);
	g_pMainWnd->m_pBtn->SetEnabled(false);

}

void CMyWebBrowserEvenrHandler::NavigateComplete2(CWebBrowserUI* pWeb, IDispatch* pDisp, VARIANT*& url)
{
	CDuiString strUrl(url->bstrVal);
	//MyOutputDebugStringW(L"NavigateComplete2:%s\r\n", strUrl);
	// 页面加载完毕才能执行js
	g_pMainWnd->m_pBtn->SetEnabled(true);

	//c++执行js示例
	// execute js start 
	IDispatch *pHtmlDocDisp = pWeb->GetHtmlWindow();
	IHTMLDocument2 *pHtmlDoc2 = NULL;
	HRESULT hr = pHtmlDocDisp->QueryInterface(IID_IHTMLDocument2, (void**)&pHtmlDoc2);
	pHtmlDocDisp->Release();
	if (SUCCEEDED(hr) && pHtmlDoc2 != NULL)
	{
		CComQIPtr<IHTMLWindow2> pHTMLWnd;
		pHtmlDoc2->get_parentWindow(&pHTMLWnd);
		if (SUCCEEDED(hr) && pHTMLWnd != NULL)
		{
			//CComBSTR bstrjs = SysAllocString(_T("document.documentElement.style.overflow = 'hidden'"));//去除水平方向滚动条  
			CComBSTR bstrjs = SysAllocString(_T("document.documentElement.style.overflowY = 'hidden'"));//去除竖直方向滚动条  
			CComBSTR bstrlan = SysAllocString(_T("javascript"));
			VARIANT varRet;
			pHTMLWnd->execScript(bstrjs, bstrlan, &varRet);
		}
	}
	// execute js end
}

void CMyWebBrowserEvenrHandler::NewWindow3(CWebBrowserUI* pWeb, IDispatch** pDisp, VARIANT_BOOL*& Cancel, DWORD dwFlags, BSTR bstrUrlContext, BSTR bstrUrl)
{
	CDuiString strUrl(bstrUrl);
	//MyOutputDebugStringW(L"NewWindow3:%s\r\n", strUrl);

}

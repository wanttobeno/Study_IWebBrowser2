#include "stdafx.h"
#include "WebBrowserExUI.h"
#include <atlcomcli.h>


int g_FunSub(int x, int y)
{
	return (x - y);
}

CWebBrowserExUI::CWebBrowserExUI()
{
}


CWebBrowserExUI::~CWebBrowserExUI()
{
}


LPCTSTR CWebBrowserExUI::GetClass() const
{
	return _T("WebBrowserExUI");
}

LPVOID CWebBrowserExUI::GetInterface(LPCTSTR pstrName)
{
	if (_tcsicmp(pstrName, _T("WebBrowserEx")) == 0)
		return static_cast<CWebBrowserExUI*>(this);

	return CActiveXUI::GetInterface(pstrName);
}


HRESULT CWebBrowserExUI::GetIDsOfNames(const IID& riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	//DISP ID 从200开始
	if (_tcscmp(rgszNames[0], _T("g_FunSub")) == 0)
		*rgDispId = 500;

	return S_OK;
}

HRESULT CWebBrowserExUI::Invoke(DISPID dispIdMember, const IID& riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	//MyOutputDebugStringW(L"%d\n", dispIdMember);
	switch (dispIdMember)
	{
	case 500:
	{
		// 注意参数顺序，反向
		VARIANTARG *varArg = pDispParams->rgvarg;
		int x = _ttoi(static_cast<_bstr_t>(varArg[1]));
		int y = _ttoi(static_cast<_bstr_t>(varArg[0]));
		int n = g_FunSub(x, y);
		*pVarResult = CComVariant(n);
		return S_OK;
	}

	default:
		break;
	}

	return CWebBrowserUI::Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

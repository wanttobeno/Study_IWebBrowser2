#pragma once

class CMyWebBrowserEvenrHandler :
	public CWebBrowserEventHandler
{
public:
	CMyWebBrowserEvenrHandler();
	~CMyWebBrowserEvenrHandler();

	virtual HRESULT STDMETHODCALLTYPE GetHostInfo(CWebBrowserUI* pWeb, /* [out][in] */ DOCHOSTUIINFO __RPC_FAR *pInfo)
	{
		if (pInfo != NULL)
		{
			pInfo->dwFlags |= DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_NO3DOUTERBORDER | DOCHOSTUIFLAG_SCROLL_NO;
		}
		return S_OK;
	}
	virtual HRESULT STDMETHODCALLTYPE ShowContextMenu(CWebBrowserUI* pWeb,
		/* [in] */ DWORD dwID,
		/* [in] */ POINT __RPC_FAR *ppt,
		/* [in] */ IUnknown __RPC_FAR *pcmdtReserved,
		/* [in] */ IDispatch __RPC_FAR *pdispReserved)
	{
		return E_NOTIMPL;
		//返回 E_NOTIMPL 正常弹出系统右键菜单
		//return S_OK;
		//返回S_OK 则可屏蔽系统右键菜单
	}
	virtual void BeforeNavigate2(CWebBrowserUI* pWeb, IDispatch *pDisp, VARIANT *&url, VARIANT *&Flags, VARIANT *&TargetFrameName, VARIANT *&PostData, VARIANT *&Headers, VARIANT_BOOL *&Cancel);
	virtual void NavigateComplete2(CWebBrowserUI* pWeb, IDispatch *pDisp, VARIANT *&url);
	virtual void NewWindow3(CWebBrowserUI* pWeb, IDispatch **pDisp, VARIANT_BOOL *&Cancel, DWORD dwFlags, BSTR bstrUrlContext, BSTR bstrUrl);
};


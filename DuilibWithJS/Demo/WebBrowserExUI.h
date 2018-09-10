#pragma once

class CWebBrowserExUI :
	public CWebBrowserUI
{
public:
	CWebBrowserExUI();
	~CWebBrowserExUI();

	LPCTSTR GetClass() const;
	LPVOID GetInterface(LPCTSTR pstrName);

	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(__RPC__in REFIID riid, __RPC__in_ecount_full(cNames) LPOLESTR *rgszNames, UINT cNames, LCID lcid, __RPC__out_ecount_full(cNames) DISPID *rgDispId);
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

	virtual HRESULT STDMETHODCALLTYPE GetExternal(IDispatch **ppDispatch)
	{
		//重写GetExternal返回一个ClientCall对象
		*ppDispatch = this;
		return S_OK;
	}
};


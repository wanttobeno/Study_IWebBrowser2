#pragma once
#include <MsHTML.h>



class WebMonitor;

class HtmlElementSink :
	public HTMLElementEvents2
{
public:
	HtmlElementSink();
	~HtmlElementSink();

	HtmlElementSink(WebMonitor* webMonitor);


	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(
		/* [out] */ __RPC__out UINT *pctinfo);

	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(
		/* [in] */ UINT iTInfo,
		/* [in] */ LCID lcid,
		/* [out] */ __RPC__deref_out_opt ITypeInfo **ppTInfo);

	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(
		/* [in] */ __RPC__in REFIID riid,
		/* [size_is][in] */ __RPC__in_ecount_full(cNames) LPOLESTR *rgszNames,
		/* [range][in] */ __RPC__in_range(0, 16384) UINT cNames,
		/* [in] */ LCID lcid,
		/* [size_is][out] */ __RPC__out_ecount_full(cNames) DISPID *rgDispId);


	STDMETHODIMP Invoke(DISPID dispidMember,
		REFIID riid,
		LCID lcid,
		WORD wFlags,
		DISPPARAMS* pdispparams,
		VARIANT* pvarResult,
		EXCEPINFO* pexcepinfo,
		UINT* puArgErr);
	//////////////////////////////

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE item(
		/* [in] */ long index,
		/* [out][retval] */ __RPC__deref_out_opt IRulesApplied **ppRulesApplied);

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_length(
		/* [out][retval] */ __RPC__out long *p);

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_element(
		/* [out][retval] */ __RPC__deref_out_opt IHTMLElement **p);

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE propertyInheritedFrom(
		/* [in] */ __RPC__in BSTR name,
		/* [out][retval] */ __RPC__deref_out_opt IRulesApplied **ppRulesApplied);

	virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_propertyCount(
		/* [out][retval] */ __RPC__out long *p);

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE property(
		/* [in] */ long index,
		/* [out][retval] */ __RPC__deref_out_opt BSTR *pbstrProperty);

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE propertyInheritedTrace(
		/* [in] */ __RPC__in BSTR name,
		/* [in] */ long index,
		/* [out][retval] */ __RPC__deref_out_opt IRulesApplied **ppRulesApplied);

	virtual /* [id] */ HRESULT STDMETHODCALLTYPE propertyInheritedTraceLength(
		/* [in] */ __RPC__in BSTR name,
		/* [out][retval] */ __RPC__out long *pLength);
	/////////////////////////


	virtual HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);

	virtual ULONG STDMETHODCALLTYPE AddRef(void);

	virtual ULONG STDMETHODCALLTYPE Release(void);
};



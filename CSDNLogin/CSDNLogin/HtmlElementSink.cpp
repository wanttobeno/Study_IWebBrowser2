#include "HtmlElementSink.h"
#include "WebMonitor.h"

#include <MsHtmdid.h> // DISPID_HTMLELEMENTEVENTS2_ONCLICK


HtmlElementSink::HtmlElementSink()
{
}


HtmlElementSink::~HtmlElementSink()
{
}


HtmlElementSink::HtmlElementSink(WebMonitor* webMonitor)
{

}

HRESULT STDMETHODCALLTYPE HtmlElementSink::GetTypeInfoCount(
	/* [out] */ __RPC__out UINT *pctinfo)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::GetTypeInfo(
	/* [in] */ UINT iTInfo,
	/* [in] */ LCID lcid,
	/* [out] */ __RPC__deref_out_opt ITypeInfo **ppTInfo)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::GetIDsOfNames(
	/* [in] */ __RPC__in REFIID riid,
	/* [size_is][in] */ __RPC__in_ecount_full(cNames) LPOLESTR *rgszNames,
	/* [range][in] */ __RPC__in_range(0, 16384) UINT cNames,
	/* [in] */ LCID lcid,
	/* [size_is][out] */ __RPC__out_ecount_full(cNames) DISPID *rgDispId)
{ 
	return S_OK;
}


STDMETHODIMP HtmlElementSink::Invoke(DISPID dispidMember,
	REFIID riid,
	LCID lcid,
	WORD wFlags,
	DISPPARAMS* pdispparams,
	VARIANT* pvarResult,
	EXCEPINFO* pexcepinfo,
	UINT* puArgErr)
{
	switch (dispidMember)
	{
	//case DISPID_HTMLELEMENTEVENTS2_ONCLICK:
	//	if (NULL != m_pEventCallBack)
	//	{
	//		m_pEventCallBack->OnClick();
	//	}
	//	break;

	default:
		break;
	}

	return S_OK;
}

/////////////////////////////////////////////////////////////////////////////
HRESULT STDMETHODCALLTYPE HtmlElementSink::item(
	/* [in] */ long index,
	/* [out][retval] */ __RPC__deref_out_opt IRulesApplied **ppRulesApplied)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::get_length(
	/* [out][retval] */ __RPC__out long *p)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::get_element(
	/* [out][retval] */ __RPC__deref_out_opt IHTMLElement **p)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::propertyInheritedFrom(
	/* [in] */ __RPC__in BSTR name,
	/* [out][retval] */ __RPC__deref_out_opt IRulesApplied **ppRulesApplied)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::get_propertyCount(
	/* [out][retval] */ __RPC__out long *p)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::property(
	/* [in] */ long index,
	/* [out][retval] */ __RPC__deref_out_opt BSTR *pbstrProperty)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::propertyInheritedTrace(
	/* [in] */ __RPC__in BSTR name,
	/* [in] */ long index,
	/* [out][retval] */ __RPC__deref_out_opt IRulesApplied **ppRulesApplied)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE HtmlElementSink::propertyInheritedTraceLength(
	/* [in] */ __RPC__in BSTR name,
	/* [out][retval] */ __RPC__out long *pLength)
{
	return S_OK;
}

///////////////////////////
HRESULT STDMETHODCALLTYPE HtmlElementSink::QueryInterface(
	/* [in] */ REFIID riid,
	/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
{
	return S_OK;
}

ULONG STDMETHODCALLTYPE HtmlElementSink::AddRef(void)
{
	return 0;
}


ULONG STDMETHODCALLTYPE HtmlElementSink::Release(void)
{
	return 0;
}
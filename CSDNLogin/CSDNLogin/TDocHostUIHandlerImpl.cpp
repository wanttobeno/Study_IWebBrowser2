#include "TDocHostUIHandlerImpl.h"


TDocHostUIHandlerImpl::TDocHostUIHandlerImpl()
{
	m_dwRef = 0;
}


TDocHostUIHandlerImpl::~TDocHostUIHandlerImpl()
{
}


HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::QueryInterface(
	/* [in] */ REFIID riid,
	/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject)
{
	//return S_OK;
	if (IsEqualIID(riid, IID_IUnknown))
	{
		*ppvObject = static_cast<IUnknown*>(this);
		return S_OK;
	}
	else if (IsEqualIID(riid, IID_IDocHostUIHandler))
	{
		*ppvObject = static_cast<IDocHostUIHandler*>(this);
		return S_OK;
	}
	else
	{
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}
}

ULONG STDMETHODCALLTYPE TDocHostUIHandlerImpl::AddRef(void)
{
	InterlockedIncrement(&m_dwRef);
	return m_dwRef;
}

ULONG STDMETHODCALLTYPE TDocHostUIHandlerImpl::Release(void)
{
	ULONG ulRefCount = InterlockedDecrement(&m_dwRef);
	if (ulRefCount <= 0)
		delete this;
	return ulRefCount;
}


/////////////////


HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::ShowContextMenu(
	/* [annotation][in] */
	_In_  DWORD dwID,
	/* [annotation][in] */
	_In_  POINT *ppt,
	/* [annotation][in] */
	_In_  IUnknown *pcmdtReserved,
	/* [annotation][in] */
	_In_  IDispatch *pdispReserved)
{
#ifdef _DEBUG
	// MessageBox(NULL, AnsiString("ShowContextMenu ID = " + IntToStr(dwID)).c_str(),
	//	NULL, MB_OK | MB_APPLMODAL | MB_ICONWARNING);
#endif // _DEBUG
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::GetHostInfo(
	/* [annotation][out][in] */
	_Inout_  DOCHOSTUIINFO *pInfo)
{
	pInfo->dwFlags = pInfo->dwFlags | DOCHOSTUIFLAG_NO3DBORDER;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::ShowUI(
	/* [annotation][in] */
	_In_  DWORD dwID,
	/* [annotation][in] */
	_In_  IOleInPlaceActiveObject *pActiveObject,
	/* [annotation][in] */
	_In_  IOleCommandTarget *pCommandTarget,
	/* [annotation][in] */
	_In_  IOleInPlaceFrame *pFrame,
	/* [annotation][in] */
	_In_  IOleInPlaceUIWindow *pDoc)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::HideUI(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::UpdateUI(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::EnableModeless(
	/* [in] */ BOOL fEnable)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::OnDocWindowActivate(
	/* [in] */ BOOL fActivate)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::OnFrameWindowActivate(
	/* [in] */ BOOL fActivate)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::ResizeBorder(
	/* [annotation][in] */
	_In_  LPCRECT prcBorder,
	/* [annotation][in] */
	_In_  IOleInPlaceUIWindow *pUIWindow,
	/* [annotation][in] */
	_In_  BOOL fRameWindow)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::TranslateAccelerator(
	/* [in] */ LPMSG lpMsg,
	/* [in] */ const GUID *pguidCmdGroup,
	/* [in] */ DWORD nCmdID)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::GetOptionKeyPath(
	/* [annotation][out] */
	_Out_  LPOLESTR *pchKey,
	/* [in] */ DWORD dw)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::GetDropTarget(
	/* [annotation][in] */
	_In_  IDropTarget *pDropTarget,
	/* [annotation][out] */
	_Outptr_  IDropTarget **ppDropTarget)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::GetExternal(
	/* [annotation][out] */
	_Outptr_result_maybenull_  IDispatch **ppDispatch)
{
	return S_OK;
}


HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::TranslateUrl(
	/* [in] */ DWORD dwTranslate,
	/* [annotation][in] */
	_In_  LPWSTR pchURLIn,
	/* [annotation][out] */
	_Outptr_  LPWSTR *ppchURLOut)
{
	return S_OK;
}


HRESULT STDMETHODCALLTYPE TDocHostUIHandlerImpl::FilterDataObject(
	/* [annotation][in] */
	_In_  IDataObject *pDO,
	/* [annotation][out] */
	_Outptr_result_maybenull_  IDataObject **ppDORet)
{
	return S_OK;
}
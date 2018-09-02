#include "WebBrowserCore.h"
#include "ConstVariables.h"
WebBrowser::WebBrowser(HWND parentWindowHandle_)
{
	
	::SetRect(&rObject, -300, -300, 300, 300);
	parentWindowHandle = parentWindowHandle_;
	if (CreateBrowser() == FALSE)
	{
		MessageBox(NULL, _T("CreateBrowser() failed"),
			_T("WebBrowser::WebBrowser"),
			MB_ICONERROR);
		return;
	}
	else {
		this->Navigate(HOME_PAGE);
	}
	
}
WebBrowser::~WebBrowser() {
	if (oleObject != nullptr) {
		delete oleObject;
	}
	if (oleInPlaceObject != nullptr) {
		delete oleInPlaceObject;
	}
	if (webBrowser2 != nullptr) {
		delete webBrowser2;
	}
	if (olePlaceActiveObject != nullptr) {
		delete olePlaceActiveObject;
	}
}

bool WebBrowser::CreateBrowser()
{
	HRESULT hr;
	hr = ::OleCreate(CLSID_WebBrowser,
		IID_IOleObject, OLERENDER_DRAW, 0, this, this,
		(void**)&oleObject);

	if (FAILED(hr))
	{
		MessageBox(NULL, _T("OleCreate failed"),
			_T("WebBrowser::CreateBrowser"),
			MB_ICONERROR);
		return FALSE;
	}

	hr = oleObject->SetClientSite(this);
	hr = OleSetContainedObject(oleObject, TRUE);

	RECT posRect;
	::SetRect(&posRect, -300, -300, 300, 300);
	hr = oleObject->DoVerb(OLEIVERB_INPLACEACTIVATE,
		NULL, this, -1, parentWindowHandle, &posRect);
	if (FAILED(hr))
	{
		MessageBox(NULL, _T("oleObject->DoVerb() failed"),
			_T("WebBrowser::CreateBrowser"),
			MB_ICONERROR);
		return FALSE;
	}

	hr = oleObject->QueryInterface(&webBrowser2);
	if (FAILED(hr))
	{
		MessageBox(NULL, 
			_T("QueryInterface get iwebbrowser2 ptr failed"),
			_T("WebBrowser::CreateBrowser"),
			MB_ICONERROR);
		return FALSE;
	}
	webBrowser2->put_Silent(VARIANT_TRUE);//ÆÁ±Îjs´íÎóÌáÊ¾
	
	webBrowser2->QueryInterface(IID_IOleInPlaceActiveObject, 
							(void**)&olePlaceActiveObject);


	return TRUE;
}

RECT WebBrowser::PixelToHiMetric(const RECT& _rc)
{
	static bool s_initialized = false;
	static int s_pixelsPerInchX, s_pixelsPerInchY;
	if(!s_initialized)
	{
		HDC hdc = ::GetDC(0);
		s_pixelsPerInchX = ::GetDeviceCaps(hdc, LOGPIXELSX);
		s_pixelsPerInchY = ::GetDeviceCaps(hdc, LOGPIXELSY);
		::ReleaseDC(0, hdc);
		s_initialized = true;
	}

	RECT rc;
	rc.left = MulDiv(2540, _rc.left, s_pixelsPerInchX);
	rc.top = MulDiv(2540, _rc.top, s_pixelsPerInchY);
	rc.right = MulDiv(2540, _rc.right, s_pixelsPerInchX);
	rc.bottom = MulDiv(2540, _rc.bottom, s_pixelsPerInchY);
	return rc;
}

void WebBrowser::SetRect(const RECT& _rc)
{
	rObject = _rc;

	{
		RECT hiMetricRect = PixelToHiMetric(rObject);
		SIZEL sz;
		sz.cx = hiMetricRect.right - hiMetricRect.left;
		sz.cy = hiMetricRect.bottom - hiMetricRect.top;
		oleObject->SetExtent(DVASPECT_CONTENT, &sz);
	}

	if(oleInPlaceObject != nullptr)
	{
		oleInPlaceObject->SetObjectRects(&rObject, &rObject);
	}
}


void WebBrowser::Navigate(wchar_t * urlPtr)
{
	variant_t flags(navNoHistory);
    this->webBrowser2->Navigate(urlPtr, &flags, 0, 0, 0);
}

void WebBrowser::TranslateMessage(LPMSG lpmsg) {
	if (olePlaceActiveObject != nullptr) {
		olePlaceActiveObject->TranslateAccelerator(lpmsg);
	}
	
}

// c++ api

void WebBrowser::AddJavaScriptFunction(wchar_t * functionName, 
	ApiToJavaScript cppFunctionPtr)
{
	JavaScriptFunction jsFunc;
	jsFunc.jsFunctionName = functionName;
	jsFunc.jsFunctionID = DISPID_BASE + javaScriptFunctions.size();
	jsFunc.cppFunctionPtr = cppFunctionPtr;
	javaScriptFunctions.push_back(jsFunc);
}





// Implement IUnknown

HRESULT STDMETHODCALLTYPE WebBrowser::QueryInterface(
	REFIID riid, void**ppvObject)
{
	if (riid == IID_IUnknown)
	{
		(*ppvObject) = static_cast<IOleClientSite*>(this);
	}
	else if (riid == IID_IOleInPlaceSite)
	{
		(*ppvObject) = static_cast<IOleInPlaceSite*>(this);
	}
	else if (riid == IID_IDispatch) {
		(*ppvObject) = static_cast<IDispatch*>(this);
	}
	else if (riid == IID_IDocHostUIHandler) {
		(*ppvObject) = static_cast<IDocHostUIHandler*>(this);
	}
	else
	{
		return E_NOINTERFACE;
	}

	AddRef();
	return S_OK;
}

ULONG STDMETHODCALLTYPE WebBrowser::AddRef(void)
{
	comReferenceCount += 1;
	return comReferenceCount;
}

ULONG STDMETHODCALLTYPE WebBrowser::Release(void)
{
	comReferenceCount -= 1;
	return comReferenceCount;
}

// Implement IOleWindow

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindow( 
	__RPC__deref_out_opt HWND *phwnd)
{
	(*phwnd) = parentWindowHandle;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ContextSensitiveHelp(
	BOOL fEnterMode)
{
	return E_NOTIMPL;
}

// Implement IOleInPlaceSite Functions

HRESULT STDMETHODCALLTYPE WebBrowser::CanInPlaceActivate(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceActivate(void)
{
	OleLockRunning(oleObject, TRUE, FALSE);
	oleObject->QueryInterface(&oleInPlaceObject);
	oleInPlaceObject->SetObjectRects(&rObject, &rObject);

	return S_OK;

}

HRESULT STDMETHODCALLTYPE WebBrowser::OnUIActivate(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindowContext( 
	__RPC__deref_out_opt IOleInPlaceFrame **ppFrame,
	__RPC__deref_out_opt IOleInPlaceUIWindow **ppDoc,
	__RPC__out LPRECT lprcPosRect,
	__RPC__out LPRECT lprcClipRect,
	__RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
	HWND hwnd = parentWindowHandle;

	(*ppFrame) = NULL;
	(*ppDoc) = NULL;
	(*lprcPosRect).left = rObject.left;
	(*lprcPosRect).top = rObject.top;
	(*lprcPosRect).right = rObject.right;
	(*lprcPosRect).bottom = rObject.bottom;
	*lprcClipRect = *lprcPosRect;

	lpFrameInfo->fMDIApp = false;
	lpFrameInfo->hwndFrame = hwnd;
	lpFrameInfo->haccel = NULL;
	lpFrameInfo->cAccelEntries = 0;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Scroll( 
	SIZE scrollExtant)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnUIDeactivate( 
	BOOL fUndoable)
{
	return S_OK;
}

HWND WebBrowser::GetControlWindow()
{
	if(controlWindowHandle != nullptr)
		return controlWindowHandle;

	if(oleInPlaceObject == nullptr)
		return 0;

	oleInPlaceObject->GetWindow(&controlWindowHandle);

	return controlWindowHandle;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceDeactivate(void)
{
	controlWindowHandle = 0;
	oleInPlaceObject = 0;

	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DiscardUndoState(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DeactivateAndUndo(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnPosRectChange( 
	__RPC__in LPCRECT lprcPosRect)
{
	return E_NOTIMPL;
}

// Implement IOleClientSite Functions

HRESULT STDMETHODCALLTYPE WebBrowser::SaveObject(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetMoniker( 
	DWORD dwAssign,
	DWORD dwWhichMoniker,
	__RPC__deref_out_opt IMoniker **ppmk)
{
	if((dwAssign == OLEGETMONIKER_ONLYIFTHERE) &&
		(dwWhichMoniker == OLEWHICHMK_CONTAINER))
		return E_FAIL;

	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetContainer( 
	__RPC__deref_out_opt IOleContainer **ppContainer)
{
	return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ShowObject(void)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnShowWindow( 
	BOOL fShow)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::RequestNewObjectLayout(void)
{
	return E_NOTIMPL;
}

// Implement IStorage Functions

HRESULT STDMETHODCALLTYPE WebBrowser::CreateStream( 
	__RPC__in_string const OLECHAR *pwcsName,
	DWORD grfMode,
	DWORD reserved1,
	DWORD reserved2,
	__RPC__deref_out_opt IStream **ppstm)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OpenStream( 
	const OLECHAR *pwcsName,
	void *reserved1,
	DWORD grfMode,
	DWORD reserved2,
	IStream **ppstm)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::CreateStorage( 
	__RPC__in_string const OLECHAR *pwcsName,
	DWORD grfMode,
	DWORD reserved1,
	DWORD reserved2,
	__RPC__deref_out_opt IStorage **ppstg)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OpenStorage( 
	__RPC__in_opt_string const OLECHAR *pwcsName,
	__RPC__in_opt IStorage *pstgPriority,
	DWORD grfMode,
	__RPC__deref_opt_in_opt SNB snbExclude,
	DWORD reserved,
	__RPC__deref_out_opt IStorage **ppstg)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::CopyTo( 
	DWORD ciidExclude,
	const IID *rgiidExclude,
	__RPC__in_opt  SNB snbExclude,
	IStorage *pstgDest)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::MoveElementTo( 
	__RPC__in_string const OLECHAR *pwcsName,
	__RPC__in_opt IStorage *pstgDest,
	__RPC__in_string const OLECHAR *pwcsNewName,
	DWORD grfFlags)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Commit( 
	DWORD grfCommitFlags)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Revert(void)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::EnumElements( 
	DWORD reserved1,
	void *reserved2,
	DWORD reserved3,
	IEnumSTATSTG **ppenum)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DestroyElement( 
	__RPC__in_string const OLECHAR *pwcsName)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::RenameElement( 
	__RPC__in_string const OLECHAR *pwcsOldName,
	__RPC__in_string const OLECHAR *pwcsNewName)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetElementTimes( 
	__RPC__in_opt_string const OLECHAR *pwcsName,
	__RPC__in_opt const FILETIME *pctime,
	__RPC__in_opt const FILETIME *patime,
	__RPC__in_opt const FILETIME *pmtime)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetClass(
	__RPC__in REFCLSID clsid)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetStateBits( 
	DWORD grfStateBits,
	DWORD grfMask)
{
	return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Stat( 
	__RPC__out STATSTG *pstatstg,
	DWORD grfStatFlag)
{
	return E_NOTIMPL;
}

// Implement IDispatch Functions

HRESULT STDMETHODCALLTYPE WebBrowser::GetTypeInfoCount(UINT *pctinfo)
{
	*pctinfo = 0;
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetTypeInfo(UINT iTInfo, LCID lcid,
	ITypeInfo **ppTInfo)
{
	return E_FAIL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetIDsOfNames(REFIID riid,
	LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
	if (cNames == 0 || rgszNames == NULL || 
		rgszNames[0] == NULL || rgDispId == NULL)
	{
		return E_INVALIDARG;
	}
	int functionCount = javaScriptFunctions.size();
	for (int i = 0; i < functionCount; i++) {
		if (_wcsicmp(rgszNames[0], javaScriptFunctions[i].jsFunctionName) == 0)
		{
			*rgDispId = javaScriptFunctions[i].jsFunctionID;
			return S_OK;
		}
	}
	return E_INVALIDARG;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Invoke(DISPID dispIdMember,
	REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *Params,
	VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
	if (dispIdMember<DISPID_BASE || dispIdMember>DISPID_BASE + javaScriptFunctions.size()) {
		return E_NOTIMPL;
	}	
	for (UINT i = 0; i < Params->cArgs; ++i)
	{
		if ((Params->rgvarg[i].vt & VT_BSTR) == 0)
		{
			return E_INVALIDARG;
		}
	}

	return javaScriptFunctions[dispIdMember-DISPID_BASE].cppFunctionPtr(dispIdMember, Params, pVarResult);
}

// Implement IDocHostUIHandler Functions

HRESULT STDMETHODCALLTYPE WebBrowser::ShowContextMenu(DWORD dwID,
	POINT *ppt, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
	
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ShowUI(
	DWORD dwID,
	IOleInPlaceActiveObject *pActiveObject, 
	IOleCommandTarget *pCommandTarget,
	IOleInPlaceFrame *pFrame,
	IOleInPlaceUIWindow *pDoc)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::HideUI()
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::UpdateUI()
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::EnableModeless(BOOL fEnable)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnDocWindowActivate(BOOL fActivate)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnFrameWindowActivate(
	BOOL fActivate)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ResizeBorder(LPCRECT prcBorder,
	IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow)
{
	return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::TranslateAccelerator(LPMSG lpMsg,
	const GUID *pguidCmdGroup, DWORD nCmdID)
{
	return S_FALSE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetOptionKeyPath(LPOLESTR *pchKey,
	DWORD dw)
{
	return S_FALSE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetDropTarget(
	IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
	return S_FALSE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetExternal(IDispatch **ppDispatch)
{
	(*ppDispatch) = static_cast<IDispatch*>(this);
	return S_OK;

}

HRESULT STDMETHODCALLTYPE WebBrowser::TranslateUrl(DWORD dwTranslate,
	OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
	*ppchURLOut = 0;

	return S_FALSE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::FilterDataObject(IDataObject *pDO,
	IDataObject **ppDORet)
{
	*ppDORet = 0;
	return S_FALSE;
}
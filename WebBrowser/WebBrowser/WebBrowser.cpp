#include "WebBrowser.h"
#include <atlbase.h>
#include <exdispid.h>


WebBrowser::WebBrowser(HWND _hWndParent)
{
	m_pOleObject = nullptr;
	m_pOleInPlaceObject = nullptr;
	m_pWebBrowser2 = nullptr;
	m_iComRefCount = 0;
	m_hWndParent = nullptr;
	m_hWndControl = nullptr;
	m_dwCookie = 0;

    m_iComRefCount = 0;
    //::SetRect(&rObject, -300, -300, 300, 300);
    m_hWndParent = _hWndParent;

    if (CreateBrowser() == FALSE)
    {
        return;
    }

    ShowWindow(GetControlWindow(), SW_SHOW);

    this->Navigate(_T("about:blank"));
}

bool WebBrowser::CreateBrowser()
{
    HRESULT hr;
    hr = ::OleCreate(CLSID_WebBrowser,
        IID_IOleObject, OLERENDER_DRAW, 0, this, this,
        (void**)&m_pOleObject);

    if (FAILED(hr))
    {
        MessageBox(NULL, _T("Cannot create oleObject CLSID_WebBrowser"),
            _T("Error"),
            MB_ICONERROR);
        return FALSE;
    }

    hr = m_pOleObject->SetClientSite(this);
    hr = OleSetContainedObject(m_pOleObject, TRUE);

    RECT posRect;
    ::GetClientRect(m_hWndParent, &posRect);
    //::SetRect(&posRect, -300, -300, 300, 300);
    hr = m_pOleObject->DoVerb(OLEIVERB_INPLACEACTIVATE,
        NULL, this, -1, m_hWndParent, &posRect);
    if (FAILED(hr))
    {
        MessageBox(NULL, _T("oleObject->DoVerb() failed"),
            _T("Error"),
            MB_ICONERROR);
        return FALSE;
    }

    hr = m_pOleObject->QueryInterface(&m_pWebBrowser2);
    if (FAILED(hr))
    {
        MessageBox(NULL, _T("oleObject->QueryInterface(&webBrowser2) failed"),
            _T("Error"),
            MB_ICONERROR);
        return FALSE;
    }

    AtlAdvise(static_cast<IUnknown*>(m_pWebBrowser2), static_cast<IDispatch*>(this), DIID_DWebBrowserEvents2, &m_dwCookie);



    return TRUE;
}

RECT WebBrowser::PixelToHiMetric(const RECT& _rc)
{
    static bool s_initialized = false;
    static int s_pixelsPerInchX, s_pixelsPerInchY;
    if (!s_initialized)
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
    m_rcObject = _rc;

    {
        RECT hiMetricRect = PixelToHiMetric(m_rcObject);
        SIZEL sz;
        sz.cx = hiMetricRect.right - hiMetricRect.left;
        sz.cy = hiMetricRect.bottom - hiMetricRect.top;
        m_pOleObject->SetExtent(DVASPECT_CONTENT, &sz);
    }

    if (m_pOleInPlaceObject != 0)
    {
        m_pOleInPlaceObject->SetObjectRects(&m_rcObject, &m_rcObject);
    }
}

// ----- Control methods -----

void WebBrowser::GoBack()
{
    this->m_pWebBrowser2->GoBack();
}

void WebBrowser::GoForward()
{
    this->m_pWebBrowser2->GoForward();
}

void WebBrowser::Refresh()
{
    this->m_pWebBrowser2->Refresh();
}

void WebBrowser::Navigate(LPCWSTR lpszUrl)
{
    bstr_t url(lpszUrl);
    variant_t flags(0x02u); //navNoHistory
    this->m_pWebBrowser2->Navigate(url, &flags, 0, 0, 0);
}

// ----- IUnknown -----

HRESULT STDMETHODCALLTYPE WebBrowser::QueryInterface(REFIID riid,
    void**ppvObject)
{
    if (riid == __uuidof(IUnknown))
    {
        (*ppvObject) = static_cast<IOleClientSite*>(this);
    }
    else if (riid == __uuidof(IOleInPlaceSite))
    {
        (*ppvObject) = static_cast<IOleInPlaceSite*>(this);
    }
    else if (riid == __uuidof(IOleCommandTarget))
    {
        (*ppvObject) = static_cast<IOleCommandTarget*>(this);
    }
    else if (riid == __uuidof(IDocHostUIHandler))
    {
        (*ppvObject) = static_cast<IDocHostUIHandler*>(this);
    }
    else if (riid == __uuidof(IDispatch))
    {
        (*ppvObject) = static_cast<IDispatch*>(this);
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
    m_iComRefCount++;
    return m_iComRefCount;
}

ULONG STDMETHODCALLTYPE WebBrowser::Release(void)
{
    m_iComRefCount--;
    return m_iComRefCount;
}

// ---------- IOleWindow ----------

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindow(
    __RPC__deref_out_opt HWND *phwnd)
{
    (*phwnd) = m_hWndParent;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ContextSensitiveHelp(
    BOOL fEnterMode)
{
    return E_NOTIMPL;
}

// ---------- IOleInPlaceSite ----------

HRESULT STDMETHODCALLTYPE WebBrowser::CanInPlaceActivate(void)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceActivate(void)
{
    OleLockRunning(m_pOleObject, TRUE, FALSE);
    m_pOleObject->QueryInterface(&m_pOleInPlaceObject);
    m_pOleInPlaceObject->SetObjectRects(&m_rcObject, &m_rcObject);

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
    HWND hwnd = m_hWndParent;

    (*ppFrame) = NULL;
    (*ppDoc) = NULL;
    (*lprcPosRect).left = m_rcObject.left;
    (*lprcPosRect).top = m_rcObject.top;
    (*lprcPosRect).right = m_rcObject.right;
    (*lprcPosRect).bottom = m_rcObject.bottom;
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
    if (m_hWndControl != 0)
        return m_hWndControl;

    if (m_pOleInPlaceObject == 0)
        return 0;

    m_pOleInPlaceObject->GetWindow(&m_hWndControl);
    return m_hWndControl;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceDeactivate(void)
{
    m_hWndControl = 0;
    m_pOleInPlaceObject = 0;

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

// ---------- IOleClientSite ----------

HRESULT STDMETHODCALLTYPE WebBrowser::SaveObject(void)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetMoniker(
    DWORD dwAssign,
    DWORD dwWhichMoniker,
    __RPC__deref_out_opt IMoniker **ppmk)
{
    if ((dwAssign == OLEGETMONIKER_ONLYIFTHERE) &&
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

// ----- IStorage -----

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


HRESULT WebBrowser::QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD* prgCmds, OLECMDTEXT *pCmdText)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut)
{
    if (!pguidCmdGroup || !IsEqualGUID(*pguidCmdGroup, CGID_DocHostCommandHandler))
        return OLECMDERR_E_UNKNOWNGROUP;

    // 不弹出查找框
    if (nCmdID == OLECMDID_FIND)
    {
        return S_OK;
    }

    else if (nCmdID == OLECMDID_SAVEAS)
    {
        // 不要弹保存对话框
        return S_OK;
    }
    else if (nCmdID == OLECMDID_SHOWSCRIPTERROR)
    {
        // 屏蔽脚本错误的对话框
        (*pvaOut).vt = VT_BOOL;
        (*pvaOut).boolVal = VARIANT_TRUE;
        return S_OK;
    }
    else if (nCmdID == OLECMDID_SHOWMESSAGE)
    {
        (*pvaOut).vt = VT_BOOL;
        (*pvaOut).boolVal = VARIANT_TRUE;
        return S_OK;
    }

    return OLECMDERR_E_NOTSUPPORTED;
}

// IDocHostUIHandler
HRESULT WebBrowser::ShowContextMenu(DWORD dwID, POINT *pptScreen, IUnknown *pcmdtReserved, IDispatch *pdispReserved)
{
    return S_OK;
}

HRESULT WebBrowser::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
    pInfo->cbSize = sizeof(DOCHOSTUIINFO);
    pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;
    pInfo->dwFlags |= DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_THEME | DOCHOSTUIFLAG_ENABLE_FORMS_AUTOCOMPLETE;
    return S_OK;
}


HRESULT WebBrowser::ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget,
    IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc)
{
    return S_OK;
}

HRESULT WebBrowser::HideUI(void)
{
    return S_OK;
}

HRESULT WebBrowser::UpdateUI(void)
{
    return S_OK;
}

HRESULT WebBrowser::EnableModeless(BOOL fEnable)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::OnDocWindowActivate(BOOL fActivate)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::OnFrameWindowActivate(BOOL fActivate)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow)
{
    return S_OK;
}

HRESULT WebBrowser::TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID)
{
    return S_FALSE;
}

HRESULT WebBrowser::GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::GetExternal(IDispatch **ppDispatch)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut)
{
    return E_NOTIMPL;
}

HRESULT WebBrowser::FilterDataObject(IDataObject *pDO, IDataObject **ppDORet)
{
    return E_NOTIMPL;
}


HRESULT WebBrowser::GetTypeInfoCount(UINT *pctinfo)
{
    return S_OK;
}
HRESULT WebBrowser::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
    return S_OK;
}

HRESULT WebBrowser::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId)
{
    return DISP_E_UNKNOWNNAME;
}

HRESULT WebBrowser::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
{
    if (dispIdMember == DISPID_VALUE && (wFlags & DISPATCH_PROPERTYGET) != 0 && pVarResult)
    {
        ::VariantInit(pVarResult);
        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = ::SysAllocString(L"[WebBrowser Object]");
        return S_OK;
    }
    else if (dispIdMember != DISPID_UNKNOWN)
    {
        if (OnInvoke(dispIdMember, wFlags, pDispParams, pVarResult))
            return S_OK;
    }
    if (pVarResult)
        pVarResult->vt = VT_EMPTY;

    return S_OK;
}

#define HANDLE_INVOKE_EVENT(id, f) case id: ATLTRACE(L#id); hRes = f(pDispParams); break;

bool WebBrowser::OnInvoke(DISPID dispIdMember, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult)
{
    HRESULT hRes = S_FALSE;

    switch (dispIdMember)
    {
        HANDLE_INVOKE_EVENT(DISPID_NEWWINDOW3, OnNewWindow3);
        // 		HANDLE_INVOKE_EVENT(DISPID_TITLECHANGE);
        // 		HANDLE_INVOKE_EVENT(DISPID_BEFORENAVIGATE2);
        // 		HANDLE_INVOKE_EVENT(DISPID_NAVIGATECOMPLETE2);
        // 		HANDLE_INVOKE_EVENT(DISPID_DOCUMENTCOMPLETE);
        // 		HANDLE_INVOKE_EVENT(DISPID_PROGRESSCHANGE);
        // 		HANDLE_INVOKE_EVENT(DISPID_COMMANDSTATECHANGE);
        // 		HANDLE_INVOKE_EVENT(DISPID_NAVIGATEERROR);
    default: break;
    }

    return hRes == S_OK;;
}

HRESULT WebBrowser::OnNewWindow3(DISPPARAMS *pDispParams)
{
    ATLASSERT(pDispParams->cArgs == 5);

    LPCWSTR lpszUrl = pDispParams->rgvarg[0].bstrVal;
    LPCWSTR lpszReferrer = pDispParams->rgvarg[1].bstrVal;
    // pDispParams->rgvarg[2].lVal
    VARIANT_BOOL *&Cancel = pDispParams->rgvarg[3].pboolVal;
    IDispatch **&ppDisp = pDispParams->rgvarg[4].ppdispVal;

    if (lpszUrl && *lpszUrl)
    {
        ::ShellExecute(nullptr, L"open", lpszUrl, nullptr, nullptr, SW_SHOW);
    }
    *Cancel = TRUE;

    ::PostMessage(m_hWndParent, WM_CLOSE, 0, 0);

    return S_OK;
}

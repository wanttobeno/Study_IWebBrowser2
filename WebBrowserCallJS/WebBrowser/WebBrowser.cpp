#include "StdAfx.h"
#include "WebBrowser.h"
#pragma comment(lib,"comctl32.lib")

#define NULLTEST_SE(fn, wstr) if((fn) == NULL) goto RETURN;
#define NULLTEST_E(fn, wstr) if((fn) == NULL) goto RETURN;
#define NULLTEST(fn) if((fn) == NULL) goto RETURN;

#define HRTEST_SE(fn, wstr)  if(FAILED(fn)) goto RETURN;
#define HRTEST_E(fn, wstr)  if(FAILED(fn)) goto RETURN;

#define RECTWIDTH(r) ((r).right-(r).left)
#define RECTHEIGHT(r) ((r).bottom-(r).top)

CWebBrowserBase::CWebBrowserBase(void):
	_refNum(0),
	_bInPlaced(false),
	_bExternalPlace(false),
	_bCalledCanInPlace(false),
	_bWebWndInited(false),
	_pOleObj(NULL), 
	_pInPlaceObj(NULL), 
	_pStorage(NULL), 
	_pWB2(NULL), 
	_pHtmlDoc2(NULL), 
	_pHtmlDoc3(NULL), 
	_pHtmlWnd2(NULL), 
	_pHtmlEvent(NULL)
{
	::memset((PVOID)&_rcWebWnd,0,sizeof(_rcWebWnd));
	HRTEST_SE(OleInitialize(0),L"Ole初始化错误");
	HRTEST_SE(StgCreateDocfile(0,STGM_READWRITE | STGM_SHARE_EXCLUSIVE | STGM_DIRECT | STGM_CREATE,0,&_pStorage),L"StgCreateDocfile错误");
	HRTEST_SE(OleCreate(CLSID_WebBrowser,IID_IOleObject,OLERENDER_DRAW,0,this,_pStorage,(void**)&_pOleObj),L"Ole创建失败");
	HRTEST_SE(_pOleObj->QueryInterface(IID_IOleInPlaceObject,(LPVOID*)&_pInPlaceObj),L"OleInPlaceObject创建失败");
RETURN:
	return;
}
CWebBrowserBase::~CWebBrowserBase(void)
{
}
//IUnknown method

STDMETHODIMP CWebBrowserBase::QueryInterface(REFIID iid,void**ppvObject)
{
	*ppvObject = NULL;

	if (iid == IID_IOleClientSite)	*ppvObject = (IOleClientSite*)this;
	else if (iid == IID_IUnknown)	*ppvObject = this;
	else if (iid == IID_IDispatch)	*ppvObject = (IDispatch*)this;
	else if (iid == IID_IDocHostUIHandler)		*ppvObject = (IDocHostUIHandler*)this;

	if ( _bExternalPlace == false)
	{
		if (iid == IID_IOleInPlaceSite)		*ppvObject = (IOleInPlaceSite*)this;		
		if (iid == IID_IOleInPlaceFrame)	*ppvObject = (IOleInPlaceFrame*)this;
		if (iid == IID_IOleInPlaceUIWindow)	*ppvObject = (IOleInPlaceUIWindow*)this;
	}

	if(*ppvObject)
	{
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CWebBrowserBase::AddRef()
{
	return ::InterlockedIncrement( &_refNum );
}

STDMETHODIMP_(ULONG) CWebBrowserBase::Release()
{
	return ::InterlockedDecrement( &_refNum );
}

// IDispatch Methods

HRESULT _stdcall CWebBrowserBase::GetTypeInfoCount(
	unsigned int * pctinfo) 
{
	return E_NOTIMPL;
}

HRESULT _stdcall CWebBrowserBase::GetTypeInfo(
	unsigned int iTInfo,
	LCID lcid,
	ITypeInfo FAR* FAR* ppTInfo) 
{
	return E_NOTIMPL;
}

HRESULT _stdcall CWebBrowserBase::GetIDsOfNames(
	REFIID riid, 
	OLECHAR FAR* FAR* rgszNames, 
	unsigned int cNames, 
	LCID lcid, 
	DISPID FAR* rgDispId 
)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CWebBrowserBase::Invoke(
	DISPID dispIdMember,
	REFIID riid,
	LCID lcid,
	WORD wFlags,
	DISPPARAMS* pDispParams,
	VARIANT* pVarResult,
	EXCEPINFO* pExcepInfo,
	unsigned int* puArgErr
)
{
	if( dispIdMember == DISPID_DOCUMENTCOMPLETE)
	{
		 OnDocumentCompleted();

		return S_OK;
	}
	if( dispIdMember == DISPID_BEFORENAVIGATE2)
	{
		return S_OK;
	}
	return E_NOTIMPL;
}

void CWebBrowserBase::OnDocumentCompleted()
{
}

//IOleClientSite methods
STDMETHODIMP CWebBrowserBase::SaveObject()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::GetMoniker(DWORD dwA,DWORD dwW,IMoniker**pm)
{
	*pm = 0;
	return E_NOTIMPL;
}
STDMETHODIMP CWebBrowserBase::GetContainer(IOleContainer**pc)
{
	*pc = 0;
	return E_FAIL;
}
STDMETHODIMP CWebBrowserBase::ShowObject()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::OnShowWindow(BOOL f)
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::RequestNewObjectLayout()
{
	return S_OK;
}

//IOleInPlaceSite
STDMETHODIMP CWebBrowserBase::GetWindow(HWND *p)
{
	*p = GetHWND();
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::ContextSensitiveHelp(BOOL)
{
	return E_NOTIMPL;
}
STDMETHODIMP CWebBrowserBase::CanInPlaceActivate()//If this function return S_FALSE, AX cannot activate in place!
{
	if ( _bInPlaced )//Does CWebBrowserBase Control already in placed?
	{
		_bCalledCanInPlace = true;
		return S_OK;
	}
	return S_FALSE;
}
STDMETHODIMP CWebBrowserBase::OnInPlaceActivate()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::OnUIActivate()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::GetWindowContext(IOleInPlaceFrame** ppFrame,IOleInPlaceUIWindow **ppDoc,LPRECT r1,LPRECT r2,LPOLEINPLACEFRAMEINFO o)
{
	*ppFrame = (IOleInPlaceFrame*)this;
	AddRef();
	*ppDoc = NULL;
	::GetClientRect(  GetHWND() ,&_rcWebWnd );
	*r1 = _rcWebWnd;
	*r2 = _rcWebWnd;
	o->cb = sizeof(OLEINPLACEFRAMEINFO);
	o->fMDIApp = false;
	o->hwndFrame = GetParent( GetHWND() );
	o->haccel = 0;
	o->cAccelEntries = 0;

	return S_OK;
}
STDMETHODIMP CWebBrowserBase::Scroll(SIZE s)
{
	return E_NOTIMPL;
}
STDMETHODIMP CWebBrowserBase::OnUIDeactivate(int)
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::OnInPlaceDeactivate()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::DiscardUndoState()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::DeactivateAndUndo()
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::OnPosRectChange(LPCRECT)
{
	return S_OK;
}

//IOleInPlaceFrame
STDMETHODIMP CWebBrowserBase::GetBorder(LPRECT l)
{
	::GetClientRect(  GetHWND() ,&_rcWebWnd );
	*l = _rcWebWnd;
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::RequestBorderSpace(LPCBORDERWIDTHS b)
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::SetBorderSpace(LPCBORDERWIDTHS b)
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::SetActiveObject(IOleInPlaceActiveObject*pV,LPCOLESTR s)
{
	return S_OK;
}
STDMETHODIMP CWebBrowserBase::SetStatusText(LPCOLESTR t)
{
	return E_NOTIMPL;
}
STDMETHODIMP CWebBrowserBase::EnableModeless(BOOL f)
{
	return E_NOTIMPL;
}
STDMETHODIMP CWebBrowserBase::TranslateAccelerator(LPMSG,WORD)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CWebBrowserBase::RemoveMenus(HMENU h)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CWebBrowserBase::InsertMenus(HMENU h,LPOLEMENUGROUPWIDTHS x)
{
	return E_NOTIMPL;
}
HRESULT _stdcall CWebBrowserBase::SetMenu(HMENU h,HOLEMENU hO,HWND hw)
{
	return E_NOTIMPL;
}

HRESULT CWebBrowserBase:: ShowContextMenu(
	 DWORD dwID,
	 POINT *ppt,
	 IUnknown *pcmdtReserved,
	 IDispatch *pdispReserved
)
{
	return E_NOTIMPL;

}
HRESULT CWebBrowserBase::GetHostInfo(DOCHOSTUIINFO *pInfo)
{
	pInfo->dwFlags |= DOCHOSTUIFLAG_NO3DOUTERBORDER;
	return S_OK;
}
HRESULT CWebBrowserBase:: ShowUI(
	DWORD dwID,
	IOleInPlaceActiveObject *pActiveObject,
	IOleCommandTarget *pCommandTarget,
	IOleInPlaceFrame *pFrame,
	IOleInPlaceUIWindow *pDoc
)
{
	return E_NOTIMPL;
}
HRESULT CWebBrowserBase::HideUI(void)
{
	return E_NOTIMPL;
}

HRESULT CWebBrowserBase::UpdateUI(void)
{
	return E_NOTIMPL;
}

HRESULT CWebBrowserBase::OnDocWindowActivate(BOOL fActivate)
{
	return E_NOTIMPL;
}
HRESULT CWebBrowserBase::OnFrameWindowActivate(BOOL fActivate)
{
	return E_NOTIMPL;
}

HRESULT CWebBrowserBase::ResizeBorder(
	LPCRECT prcBorder,
	IOleInPlaceUIWindow *pUIWindow,
	BOOL fRameWindow
)
{
	return E_NOTIMPL;
}
HRESULT CWebBrowserBase::TranslateAccelerator( 
	LPMSG lpMsg,
	const GUID *pguidCmdGroup,
	DWORD nCmdID
)
{
	return E_NOTIMPL;
}
HRESULT CWebBrowserBase::GetOptionKeyPath(LPOLESTR *pchKey,DWORD dw)
{
	return E_NOTIMPL;
}
HRESULT CWebBrowserBase::GetDropTarget(IDropTarget *pDropTarget,IDropTarget **ppDropTarget)
{
	//S_OK:自定义拖拽
	//E_NOTIMPL:使用默认拖拽
	return E_NOTIMPL;
	
}

HRESULT CWebBrowserBase::GetExternal( IDispatch **ppDispatch)
{
	return E_NOTIMPL;
}

HRESULT CWebBrowserBase::TranslateUrl(
	DWORD dwTranslate,
	OLECHAR *pchURLIn,
	OLECHAR **ppchURLOut
)
{
	return E_NOTIMPL;
}
HRESULT CWebBrowserBase:: FilterDataObject(IDataObject *pDO,IDataObject **ppDORet)
{
	return E_NOTIMPL;
}

//Other Methods
IWebBrowser2* CWebBrowserBase::GetWebBrowser2()
{
	if(_pWB2 != NULL) return _pWB2;
	NULLTEST_SE(_pOleObj,L"Ole对象为空");
	HRTEST_SE(_pOleObj->QueryInterface(IID_IWebBrowser2,(void**)&_pWB2),L"QueryInterface IID_IWebBrowser2 失败");
	return _pWB2;
RETURN:
	return NULL;
}

IDispatch* CWebBrowserBase::GetHtmlWindow()
{
	if(_pHtmlWnd2 != NULL) return _pHtmlWnd2;

	IWebBrowser2* pWB2 = GetWebBrowser2();
	NULLTEST(pWB2);

	IHTMLDocument2 *pHtmlDoc2 = NULL;
	IDispatch* pDp =  NULL;
	HRTEST_SE(pWB2->get_Document(&pDp),L"DWebBrowser2::get_Document 错误");
	HRTEST_SE(pDp->QueryInterface(IID_IHTMLDocument2,(void**)&pHtmlDoc2),L"QueryInterface IID_IHTMLDocument2 失败");

	HRTEST_SE(pHtmlDoc2->get_parentWindow(&_pHtmlWnd2),L"IHTMLWindow2::get_parentWindow 错误");

	IDispatch *pHtmlWindown = NULL;
	HRTEST_SE(_pHtmlWnd2->QueryInterface(IID_IDispatch, (void**)&pHtmlWindown),L"GetDispHTMLWindow2错误");

	pHtmlDoc2->Release();

	return pHtmlWindown;

RETURN:
	return NULL;
}

BOOL CWebBrowserBase::SetWebRect(LPRECT lprc)
{
	BOOL bRet = FALSE;
	if( false == _bInPlaced )//尚未OpenWebBrowser操作,直接写入_rcWebWnd
	{
		_rcWebWnd = *lprc;
	}
	else//已经打开CWebBrowserBase,通过 IOleInPlaceObject::SetObjectRects 调整大小
	{
		SIZEL size;
		size.cx = RECTWIDTH(*lprc);
		size.cy = RECTHEIGHT(*lprc);
		IOleObject* pOleObj;
		NULLTEST( pOleObj= _GetOleObject());
		HRTEST_E( pOleObj->SetExtent(  1,&size ),L"SetExtent 错误");
		IOleInPlaceObject* pInPlace;
		NULLTEST( pInPlace = _GetInPlaceObject());
		HRTEST_E( pInPlace->SetObjectRects(lprc,lprc),L"SetObjectRects 错误");
		_rcWebWnd = *lprc;
	}
	bRet = TRUE;
RETURN:
	return bRet;
}

BOOL CWebBrowserBase::OpenWebBrowser()
{
	BOOL bRet = FALSE;
	NULLTEST_E(_GetOleObject(),L"ActiveX对象为空" );//对于本身的实现函数,其自身承担错误录入工作

	if((RECTWIDTH(_rcWebWnd) && RECTHEIGHT(_rcWebWnd)) == 0)
	{
		::GetClientRect( GetHWND() ,&_rcWebWnd);//设置CWebBrowserBase的大小为窗口的客户区大小.
	}

	if( _bInPlaced == false )// Activate In Place
	{
		_bInPlaced = true;//_bInPlaced must be set as true, before INPLACEACTIVATE, otherwise, once DoVerb, it would return error;
		_bExternalPlace = 0;//lParam;

		HRTEST_E( _GetOleObject()->DoVerb(OLEIVERB_INPLACEACTIVATE,0,this,0, GetHWND()  ,&_rcWebWnd),L"关于INPLACE的DoVerb错误");
		_bInPlaced = true;

		//* 挂接DWebBrwoser2Event
		IConnectionPointContainer* pCPC = NULL;
		IConnectionPoint* pCP  = NULL;
		DWORD dwCookie = 0;
		HRTEST_E(GetWebBrowser2()->QueryInterface(IID_IConnectionPointContainer,(void**)&pCPC),L"枚举IConnectionPointContainer接口失败");
		HRTEST_E(pCPC->FindConnectionPoint(DIID_DWebBrowserEvents2,&pCP),L"FindConnectionPoint失败");
		HRTEST_E( pCP->Advise( (IUnknown*)(void*)this,&dwCookie),L"IConnectionPoint::Advise失败");
	}
	bRet = TRUE;
RETURN:
	return bRet;
}

BOOL CWebBrowserBase::OpenURL(VARIANT* pVarUrl)
{
	BOOL bRet = FALSE;
	HRTEST_E( GetWebBrowser2()->Navigate2( pVarUrl,0,0,0,0),L"GetWebBrowser2 失败");
	bRet = TRUE;
RETURN:
	return bRet;
}

DISPID CWebBrowserBase::FindId(IDispatch *pObj, LPOLESTR pName)
{
	DISPID id = 0;
	if(FAILED(pObj->GetIDsOfNames(IID_NULL,&pName,1,LOCALE_SYSTEM_DEFAULT,&id))) id = -1;
	return id;
}

HRESULT CWebBrowserBase::InvokeMethod(IDispatch *pObj, LPOLESTR pName, VARIANT *pVarResult, VARIANT *p, int cArgs)
{
	DISPID dispid = FindId(pObj, pName);
	if(dispid == -1) return E_FAIL;

	DISPPARAMS ps;
	ps.cArgs = cArgs;
	ps.rgvarg = p;
	ps.cNamedArgs = 0;
	ps.rgdispidNamedArgs = NULL;

	return pObj->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &ps, pVarResult, NULL, NULL);
}

HRESULT CWebBrowserBase::GetProperty(IDispatch *pObj, LPOLESTR pName, VARIANT *pValue)
{
	DISPID dispid = FindId(pObj, pName);
	if(dispid == -1) return E_FAIL;

	DISPPARAMS ps;
	ps.cArgs = 0;
	ps.rgvarg = NULL;
	ps.cNamedArgs = 0;
	ps.rgdispidNamedArgs = NULL;

	return pObj->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &ps, pValue, NULL, NULL);
}

HRESULT CWebBrowserBase::SetProperty(IDispatch *pObj, LPOLESTR pName, VARIANT *pValue)
{
	DISPID dispid = FindId(pObj, pName);
	if(dispid == -1) return E_FAIL;

	DISPPARAMS ps;
	ps.cArgs = 1;
	ps.rgvarg = pValue;
	ps.cNamedArgs = 0;
	ps.rgdispidNamedArgs = NULL;

	return pObj->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYPUT, &ps, NULL, NULL, NULL);
}


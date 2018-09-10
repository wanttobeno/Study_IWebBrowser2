#pragma once

#include <mshtmhst.h>
#include <mshtml.h>
#include <Exdisp.h>
#include <exdispid.h>

template<class T> class CComPtr
{
private:
	T ptr;
public:
	CComPtr()
	{ 
		ptr = NULL;
	}
	CComPtr(T p)
	{ 
		ptr = p;
	}

	~CComPtr()
	{ 
		if(ptr != NULL) ptr->Release(); 
	}

	operator T()
	{
		return ptr;
	}

	T operator->()
	{
		return ptr;
	}

	T* operator&()
	{
		return &ptr;
	}
};

class CVariant : public VARIANT
{
public:
	CVariant() 
	{ 
		VariantInit(this); 
	}
	CVariant(int i)
	{
		VariantInit(this);
		this->vt = VT_I4;
		this->intVal = i;
	}
	CVariant(float f)
	{
		VariantInit(this);
		this->vt = VT_R4;
		this->fltVal = f;
	}
	CVariant(LPOLESTR s)
	{
		VariantInit(this);
		this->vt = VT_BSTR;
		this->bstrVal = s;
	}
	CVariant(IDispatch *disp)
	{
		VariantInit(this);
		this->vt = VT_DISPATCH;
		this->pdispVal = disp;
	}

	~CVariant() 
	{ 
		VariantClear(this); 
	}
};

class CWebBrowserBase:
	public IDispatch,
	public IOleClientSite,
	public IOleInPlaceSite,
	public IOleInPlaceFrame,
	public IDocHostUIHandler
{
public:
	CWebBrowserBase();
	~CWebBrowserBase(void);
public:
	// IUnknown methods
	virtual STDMETHODIMP QueryInterface(REFIID iid,void**ppvObject);
	virtual STDMETHODIMP_(ULONG) AddRef();
	virtual STDMETHODIMP_(ULONG) Release();
	// IDispatch Methods
	HRESULT _stdcall GetTypeInfoCount(unsigned int * pctinfo);
	HRESULT _stdcall GetTypeInfo(unsigned int iTInfo,LCID lcid,ITypeInfo FAR* FAR* ppTInfo);
	HRESULT _stdcall GetIDsOfNames(REFIID riid,OLECHAR FAR* FAR* rgszNames,unsigned int cNames,LCID lcid,DISPID FAR* rgDispId);
	HRESULT _stdcall Invoke(DISPID dispIdMember,REFIID riid,LCID lcid,WORD wFlags,DISPPARAMS FAR* pDispParams,VARIANT FAR* pVarResult,EXCEPINFO FAR* pExcepInfo,unsigned int FAR* puArgErr);
	// IOleClientSite methods
	virtual STDMETHODIMP SaveObject();
	virtual STDMETHODIMP GetMoniker(DWORD dwA,DWORD dwW,IMoniker**pm);
	virtual STDMETHODIMP GetContainer(IOleContainer**pc);
	virtual STDMETHODIMP ShowObject();
	virtual STDMETHODIMP OnShowWindow(BOOL f);
	virtual STDMETHODIMP RequestNewObjectLayout();
	// IOleInPlaceSite methods
	virtual STDMETHODIMP GetWindow(HWND *p);
	virtual STDMETHODIMP ContextSensitiveHelp(BOOL);
	virtual STDMETHODIMP CanInPlaceActivate();
	virtual STDMETHODIMP OnInPlaceActivate();
	virtual STDMETHODIMP OnUIActivate();
	virtual STDMETHODIMP GetWindowContext(IOleInPlaceFrame** ppFrame,IOleInPlaceUIWindow **ppDoc,LPRECT r1,LPRECT r2,LPOLEINPLACEFRAMEINFO o);
	virtual STDMETHODIMP Scroll(SIZE s);
	virtual STDMETHODIMP OnUIDeactivate(int);
	virtual STDMETHODIMP OnInPlaceDeactivate();
	virtual STDMETHODIMP DiscardUndoState();
	virtual STDMETHODIMP DeactivateAndUndo();
	virtual STDMETHODIMP OnPosRectChange(LPCRECT);
	// IOleInPlaceFrame methods
	virtual STDMETHODIMP GetBorder(LPRECT l);
	virtual STDMETHODIMP RequestBorderSpace(LPCBORDERWIDTHS);
	virtual STDMETHODIMP SetBorderSpace(LPCBORDERWIDTHS w);
	virtual STDMETHODIMP SetActiveObject(IOleInPlaceActiveObject*pV,LPCOLESTR s);
	virtual STDMETHODIMP InsertMenus(HMENU h,LPOLEMENUGROUPWIDTHS x);
	virtual STDMETHODIMP SetMenu(HMENU h,HOLEMENU hO,HWND hw);
	virtual STDMETHODIMP RemoveMenus(HMENU h);
	virtual STDMETHODIMP SetStatusText(LPCOLESTR t);
	virtual STDMETHODIMP TranslateAccelerator(LPMSG,WORD);
	//IDocHostUIHandler
	virtual HRESULT STDMETHODCALLTYPE ShowContextMenu(DWORD dwID,POINT *ppt,IUnknown *pcmdtReserved,IDispatch *pdispReserved);
	virtual HRESULT STDMETHODCALLTYPE GetHostInfo(DOCHOSTUIINFO *pInfo);
	virtual HRESULT STDMETHODCALLTYPE ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame,IOleInPlaceUIWindow *pDoc);
	virtual HRESULT STDMETHODCALLTYPE HideUI( void);
	virtual HRESULT STDMETHODCALLTYPE UpdateUI( void);
	virtual HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable);
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL fActivate);
	virtual HRESULT STDMETHODCALLTYPE OnFrameWindowActivate( BOOL fActivate);
	virtual HRESULT STDMETHODCALLTYPE ResizeBorder(LPCRECT prcBorder,IOleInPlaceUIWindow *pUIWindow,BOOL fRameWindow);
	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpMsg,const GUID *pguidCmdGroup,DWORD nCmdID);
	virtual HRESULT STDMETHODCALLTYPE GetOptionKeyPath(LPOLESTR *pchKey,DWORD dw);
	virtual HRESULT STDMETHODCALLTYPE GetDropTarget(IDropTarget *pDropTarget,IDropTarget **ppDropTarget);
	virtual HRESULT STDMETHODCALLTYPE GetExternal( IDispatch **ppDispatch);
	virtual HRESULT STDMETHODCALLTYPE TranslateUrl(DWORD dwTranslate,OLECHAR *pchURLIn,OLECHAR **ppchURLOut);
	virtual HRESULT STDMETHODCALLTYPE FilterDataObject(IDataObject *pDO,IDataObject **ppDORet);

protected:
	virtual HWND GetHWND(){return NULL;}
	// 内部工具函数
private:
	inline IOleObject* _GetOleObject(){return _pOleObj;};
	inline IOleInPlaceObject* _GetInPlaceObject(){return _pInPlaceObj;};
	//外部方法
protected:
	virtual void OnDocumentCompleted();
public:
	IWebBrowser2*      GetWebBrowser2();
	IDispatch*		   GetHtmlWindow();

	BOOL SetWebRect(LPRECT lprc);
	BOOL OpenWebBrowser();
	BOOL OpenURL(VARIANT* pVarUrl);

	static DISPID FindId(IDispatch *pObj, LPOLESTR pName);
	static HRESULT InvokeMethod(IDispatch *pObj, LPOLESTR pMehtod, VARIANT *pVarResult, VARIANT *ps, int cArgs);
	static HRESULT GetProperty(IDispatch *pObj, LPOLESTR pName, VARIANT *pValue);
	static HRESULT SetProperty(IDispatch *pObj, LPOLESTR pName, VARIANT *pValue);

	// 内部数据
protected:
	long   _refNum;
private:
	RECT  _rcWebWnd;
	bool  _bInPlaced;
	bool  _bExternalPlace;
	bool  _bCalledCanInPlace;
	bool  _bWebWndInited;
private:
	//指针
	IOleObject*                 _pOleObj; 
	IOleInPlaceObject*          _pInPlaceObj;
	IStorage*                   _pStorage;
	IWebBrowser2*               _pWB2;
	IHTMLDocument2*             _pHtmlDoc2;
	IHTMLDocument3*             _pHtmlDoc3;
	IHTMLWindow2*               _pHtmlWnd2;
	IHTMLEventObj*              _pHtmlEvent;
};

#pragma once
#include <comdef.h>
#include <Exdisp.h>
#include <string>
#include <tchar.h>
#include <mshtmhst.h>
#include <vector>
#include "JavaScriptFunction.h"


using namespace std;

class WebBrowser :
	public IOleClientSite,
	public IOleInPlaceSite,
	public IStorage,
	public IDispatch,
	public IDocHostUIHandler
{
public:

	WebBrowser(HWND hWndParent);
	~WebBrowser();

	bool CreateBrowser();

	RECT PixelToHiMetric(const RECT& _rc);

	virtual void SetRect(const RECT& _rc);

	// Tool Functions

	
	void Navigate(wchar_t * urlPtr);

	void TranslateMessage(LPMSG lpmsg);

	// c++ api to javascript

	void AddJavaScriptFunction(wchar_t * functionName,
		ApiToJavaScript cppFunctionPtr);


	// IUnknown Virtual Functions

	virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
		void**ppvObject) override;

	virtual ULONG STDMETHODCALLTYPE AddRef(void);

	virtual ULONG STDMETHODCALLTYPE Release(void);

	// IOleWindow Virtual Functions

	virtual HRESULT STDMETHODCALLTYPE GetWindow( 
		__RPC__deref_out_opt HWND *phwnd) override;

	virtual HRESULT STDMETHODCALLTYPE ContextSensitiveHelp( 
		BOOL fEnterMode) override;

	// IOleInPlaceSite Virtual Functions

	virtual HRESULT STDMETHODCALLTYPE CanInPlaceActivate(void) override;

	virtual HRESULT STDMETHODCALLTYPE OnInPlaceActivate(void) override;

	virtual HRESULT STDMETHODCALLTYPE OnUIActivate(void) override;

	virtual HRESULT STDMETHODCALLTYPE GetWindowContext( 
		__RPC__deref_out_opt IOleInPlaceFrame **ppFrame,
		__RPC__deref_out_opt IOleInPlaceUIWindow **ppDoc,
		__RPC__out LPRECT lprcPosRect,
		__RPC__out LPRECT lprcClipRect,
		__RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo) override;

	virtual HRESULT STDMETHODCALLTYPE Scroll( 
		SIZE scrollExtant) override;

	virtual HRESULT STDMETHODCALLTYPE OnUIDeactivate( 
		BOOL fUndoable) override;

	virtual HWND GetControlWindow();

	virtual HRESULT STDMETHODCALLTYPE OnInPlaceDeactivate(void) override;

	virtual HRESULT STDMETHODCALLTYPE DiscardUndoState(void) override;

	virtual HRESULT STDMETHODCALLTYPE DeactivateAndUndo(void) override;

	virtual HRESULT STDMETHODCALLTYPE OnPosRectChange( 
		__RPC__in LPCRECT lprcPosRect) override;

	// IOleClientSite Virtual Functions

	virtual HRESULT STDMETHODCALLTYPE SaveObject(void) override;

	virtual HRESULT STDMETHODCALLTYPE GetMoniker( 
		DWORD dwAssign,
		DWORD dwWhichMoniker,
		__RPC__deref_out_opt IMoniker **ppmk) override;

	virtual HRESULT STDMETHODCALLTYPE GetContainer( 
		__RPC__deref_out_opt IOleContainer **ppContainer) override;

	virtual HRESULT STDMETHODCALLTYPE ShowObject(void) override;
	virtual HRESULT STDMETHODCALLTYPE OnShowWindow( 
		BOOL fShow) override;

	virtual HRESULT STDMETHODCALLTYPE RequestNewObjectLayout(void) override;

	// IStorage Virtual Functions

	virtual HRESULT STDMETHODCALLTYPE CreateStream( 
		__RPC__in_string const OLECHAR *pwcsName,
		DWORD grfMode,
		DWORD reserved1,
		DWORD reserved2,
		__RPC__deref_out_opt IStream **ppstm) override;

	virtual HRESULT STDMETHODCALLTYPE OpenStream( 
		const OLECHAR *pwcsName,
		void *reserved1,
		DWORD grfMode,
		DWORD reserved2,
		IStream **ppstm) override;

	virtual HRESULT STDMETHODCALLTYPE CreateStorage( 
		__RPC__in_string const OLECHAR *pwcsName,
		DWORD grfMode,
		DWORD reserved1,
		DWORD reserved2,
		__RPC__deref_out_opt IStorage **ppstg) override;

	virtual HRESULT STDMETHODCALLTYPE OpenStorage(
		__RPC__in_opt_string const OLECHAR *pwcsName,
		__RPC__in_opt IStorage *pstgPriority,
		DWORD grfMode,
		__RPC__deref_opt_in_opt SNB snbExclude,
		DWORD reserved,
		__RPC__deref_out_opt IStorage **ppstg) override;

	virtual HRESULT STDMETHODCALLTYPE CopyTo( 
		DWORD ciidExclude,
		const IID *rgiidExclude,
		__RPC__in_opt  SNB snbExclude,
		IStorage *pstgDest) override;

	virtual HRESULT STDMETHODCALLTYPE MoveElementTo( 
		__RPC__in_string const OLECHAR *pwcsName,
		__RPC__in_opt IStorage *pstgDest,
		__RPC__in_string const OLECHAR *pwcsNewName,
		DWORD grfFlags) override;

	virtual HRESULT STDMETHODCALLTYPE Commit( 
		DWORD grfCommitFlags) override;

	virtual HRESULT STDMETHODCALLTYPE Revert(void) override;

	virtual HRESULT STDMETHODCALLTYPE EnumElements( 
		DWORD reserved1,
		void *reserved2,
		DWORD reserved3,
		IEnumSTATSTG **ppenum) override;

	virtual HRESULT STDMETHODCALLTYPE DestroyElement( 
		__RPC__in_string const OLECHAR *pwcsName) override;

	virtual HRESULT STDMETHODCALLTYPE RenameElement( 
		__RPC__in_string const OLECHAR *pwcsOldName,
		__RPC__in_string const OLECHAR *pwcsNewName) override;

	virtual HRESULT STDMETHODCALLTYPE SetElementTimes( 
		__RPC__in_opt_string const OLECHAR *pwcsName,
		__RPC__in_opt const FILETIME *pctime,
		__RPC__in_opt const FILETIME *patime,
		__RPC__in_opt const FILETIME *pmtime) override;

	virtual HRESULT STDMETHODCALLTYPE SetClass( 
		__RPC__in REFCLSID clsid) override;
	virtual HRESULT STDMETHODCALLTYPE SetStateBits( 
		DWORD grfStateBits,
		DWORD grfMask) override;

	virtual HRESULT STDMETHODCALLTYPE Stat( 
		__RPC__out STATSTG *pstatstg,
		DWORD grfStatFlag) override;

	// IDispatch Virtual Functions
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo) override;
	virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid,
		ITypeInfo **ppTInfo) override;
	virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid,
		LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId) override;
	virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid,
		LCID lcid, WORD wFlags, DISPPARAMS *Params, VARIANT *pVarResult,
		EXCEPINFO *pExcepInfo, UINT *puArgErr) override;

	// IDocHostUIHandler Virtual Functions
	virtual HRESULT STDMETHODCALLTYPE ShowContextMenu(DWORD dwID, POINT *ppt,
		IUnknown *pcmdtReserved, IDispatch *pdispReserved) override;
	virtual HRESULT STDMETHODCALLTYPE GetHostInfo(DOCHOSTUIINFO *pInfo) override;
	virtual HRESULT STDMETHODCALLTYPE ShowUI(DWORD dwID,
		IOleInPlaceActiveObject *pActiveObject,
		IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame,
		IOleInPlaceUIWindow *pDoc) override;
	virtual HRESULT STDMETHODCALLTYPE HideUI() override;
	virtual HRESULT STDMETHODCALLTYPE UpdateUI() override;
	virtual HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable) override;
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL fActivate) 
		override;
	virtual HRESULT STDMETHODCALLTYPE OnFrameWindowActivate(BOOL fActivate) 
		override;
	virtual HRESULT STDMETHODCALLTYPE ResizeBorder(LPCRECT prcBorder,
		IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow) override;
	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpMsg,
		const GUID *pguidCmdGroup, DWORD nCmdID) override;
	virtual HRESULT STDMETHODCALLTYPE GetOptionKeyPath(LPOLESTR *pchKey,
		DWORD dw) override;
	virtual HRESULT STDMETHODCALLTYPE GetDropTarget(IDropTarget *pDropTarget,
		IDropTarget **ppDropTarget) override;
	virtual HRESULT STDMETHODCALLTYPE GetExternal(IDispatch **ppDispatch) 
		override;
	virtual HRESULT STDMETHODCALLTYPE TranslateUrl(DWORD dwTranslate,
		OLECHAR *pchURLIn, OLECHAR **ppchURLOut) override;
	virtual HRESULT STDMETHODCALLTYPE FilterDataObject(IDataObject *pDO,
		IDataObject **ppDORet) override;


protected:
	IOleObject* oleObject = nullptr;
	IOleInPlaceObject* oleInPlaceObject = nullptr;
	IOleInPlaceActiveObject* olePlaceActiveObject = nullptr;

	IWebBrowser2* webBrowser2 = nullptr;

	LONG comReferenceCount = 0;

	RECT rObject;

	std::vector<JavaScriptFunction> javaScriptFunctions;

	HWND parentWindowHandle = nullptr;
	HWND controlWindowHandle = nullptr;
};

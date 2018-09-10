#pragma once

#include <comdef.h>
#include <Exdisp.h>
#include <tchar.h>
#include <Windows.h>
#include <mshtmhst.h>
#ifndef nullptr
#define nullptr NULL
#endif 

class WebBrowser :
    public IOleClientSite,
    public IOleInPlaceSite,
    public IStorage,
    public IOleCommandTarget,
    public IDocHostUIHandler,
    public IDispatch
{
public:
    WebBrowser(HWND hWndParent);

    bool CreateBrowser();

    RECT PixelToHiMetric(const RECT& _rc);

    virtual void SetRect(const RECT& _rc);

    // ----- Control methods -----

    void GoBack();

    void GoForward();

    void Refresh();

    void Navigate(LPCWSTR lpszUrl);

    // ----- IUnknown -----

    virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,
        void**ppvObject) override;

    virtual ULONG STDMETHODCALLTYPE AddRef(void);

    virtual ULONG STDMETHODCALLTYPE Release(void);

    // ---------- IOleWindow ----------

    virtual HRESULT STDMETHODCALLTYPE GetWindow(
        __RPC__deref_out_opt HWND *phwnd) override;

    virtual HRESULT STDMETHODCALLTYPE ContextSensitiveHelp(
        BOOL fEnterMode) override;

    // ---------- IOleInPlaceSite ----------

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

    // ---------- IOleClientSite ----------

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

    // ----- IStorage -----

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

    //IOleCommandTarget
    STDMETHOD(QueryStatus)(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD* prgCmds, OLECMDTEXT *pCmdText);
    STDMETHOD(Exec)(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvaIn, VARIANT *pvaOut);

    //IDocHostUIHandler
    HRESULT STDMETHODCALLTYPE ShowContextMenu(DWORD dwID, POINT *pptScreen, IUnknown *pcmdtReserved, IDispatch *pdispReserved);
    HRESULT STDMETHODCALLTYPE GetHostInfo(DOCHOSTUIINFO *pInfo);
    HRESULT STDMETHODCALLTYPE ShowUI(DWORD dwID, IOleInPlaceActiveObject *pActiveObject, IOleCommandTarget *pCommandTarget, IOleInPlaceFrame *pFrame, IOleInPlaceUIWindow *pDoc);
    HRESULT STDMETHODCALLTYPE HideUI(void);
    HRESULT STDMETHODCALLTYPE UpdateUI(void);
    HRESULT STDMETHODCALLTYPE EnableModeless(BOOL fEnable);
    HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL fActivate);
    HRESULT STDMETHODCALLTYPE OnFrameWindowActivate(BOOL fActivate);
    HRESULT STDMETHODCALLTYPE ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow *pUIWindow, BOOL fRameWindow);
    HRESULT STDMETHODCALLTYPE TranslateAccelerator(LPMSG lpMsg, const GUID *pguidCmdGroup, DWORD nCmdID);
    HRESULT STDMETHODCALLTYPE GetOptionKeyPath(LPOLESTR *pchKey, DWORD dw);
    HRESULT STDMETHODCALLTYPE GetDropTarget(IDropTarget *pDropTarget, IDropTarget **ppDropTarget);
    HRESULT STDMETHODCALLTYPE GetExternal(IDispatch **ppDispatch);
    HRESULT STDMETHODCALLTYPE TranslateUrl(DWORD dwTranslate, OLECHAR *pchURLIn, OLECHAR **ppchURLOut);
    HRESULT STDMETHODCALLTYPE FilterDataObject(IDataObject *pDO, IDataObject **ppDORet);

    // IDispatch
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo);
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId);
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr);

private:
    bool OnInvoke(DISPID dispIdMember, WORD wFlags, DISPPARAMS *pDispParams, VARIANT *pVarResult);
    HRESULT OnNewWindow3(DISPPARAMS *pDispParams);

protected:
    IOleObject * m_pOleObject;
    IOleInPlaceObject* m_pOleInPlaceObject;

    IWebBrowser2* m_pWebBrowser2;

    LONG m_iComRefCount;

    RECT m_rcObject;

    HWND m_hWndParent;
    HWND m_hWndControl;

    DWORD m_dwCookie;
};

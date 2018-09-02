/*
	IDocHostUIHandler 是个纯虚类，这里只是重写父类的接口,没有代码实现 
*/

#pragma once
#include <MsHtmHst.h>
class TDocHostUIHandlerImpl :
	public IDocHostUIHandler
{
public:
	TDocHostUIHandlerImpl();
	~TDocHostUIHandlerImpl();

	virtual HRESULT STDMETHODCALLTYPE QueryInterface(
		/* [in] */ REFIID riid,
		/* [iid_is][out] */ _COM_Outptr_ void __RPC_FAR *__RPC_FAR *ppvObject);

	virtual ULONG STDMETHODCALLTYPE AddRef(void);

	virtual ULONG STDMETHODCALLTYPE Release(void);


	/////////////////////

	virtual HRESULT STDMETHODCALLTYPE ShowContextMenu(
		/* [annotation][in] */
		_In_  DWORD dwID,
		/* [annotation][in] */
		_In_  POINT *ppt,
		/* [annotation][in] */
		_In_  IUnknown *pcmdtReserved,
		/* [annotation][in] */
		_In_  IDispatch *pdispReserved);

	virtual HRESULT STDMETHODCALLTYPE GetHostInfo(
		/* [annotation][out][in] */
		_Inout_  DOCHOSTUIINFO *pInfo);

	virtual HRESULT STDMETHODCALLTYPE ShowUI(
		/* [annotation][in] */
		_In_  DWORD dwID,
		/* [annotation][in] */
		_In_  IOleInPlaceActiveObject *pActiveObject,
		/* [annotation][in] */
		_In_  IOleCommandTarget *pCommandTarget,
		/* [annotation][in] */
		_In_  IOleInPlaceFrame *pFrame,
		/* [annotation][in] */
		_In_  IOleInPlaceUIWindow *pDoc);

	virtual HRESULT STDMETHODCALLTYPE HideUI(void);

	virtual HRESULT STDMETHODCALLTYPE UpdateUI(void);

	virtual HRESULT STDMETHODCALLTYPE EnableModeless(
		/* [in] */ BOOL fEnable);

	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(
		/* [in] */ BOOL fActivate);

	virtual HRESULT STDMETHODCALLTYPE OnFrameWindowActivate(
		/* [in] */ BOOL fActivate);

	virtual HRESULT STDMETHODCALLTYPE ResizeBorder(
		/* [annotation][in] */
		_In_  LPCRECT prcBorder,
		/* [annotation][in] */
		_In_  IOleInPlaceUIWindow *pUIWindow,
		/* [annotation][in] */
		_In_  BOOL fRameWindow);

	virtual HRESULT STDMETHODCALLTYPE TranslateAccelerator(
		/* [in] */ LPMSG lpMsg,
		/* [in] */ const GUID *pguidCmdGroup,
		/* [in] */ DWORD nCmdID);

	virtual HRESULT STDMETHODCALLTYPE GetOptionKeyPath(
		/* [annotation][out] */
		_Out_  LPOLESTR *pchKey,
		/* [in] */ DWORD dw);

	virtual HRESULT STDMETHODCALLTYPE GetDropTarget(
		/* [annotation][in] */
		_In_  IDropTarget *pDropTarget,
		/* [annotation][out] */
		_Outptr_  IDropTarget **ppDropTarget);

	virtual HRESULT STDMETHODCALLTYPE GetExternal(
		/* [annotation][out] */
		_Outptr_result_maybenull_  IDispatch **ppDispatch);

	virtual HRESULT STDMETHODCALLTYPE TranslateUrl(
		/* [in] */ DWORD dwTranslate,
		/* [annotation][in] */
		_In_  LPWSTR pchURLIn,
		/* [annotation][out] */
		_Outptr_  LPWSTR *ppchURLOut);

	virtual HRESULT STDMETHODCALLTYPE FilterDataObject(
		/* [annotation][in] */
		_In_  IDataObject *pDO,
		/* [annotation][out] */
		_Outptr_result_maybenull_  IDataObject **ppDORet);

private:
	LONG m_dwRef;
};


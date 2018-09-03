
#include "ExDispid.h"
#ifndef _ATL_MAX_VARTYPES
#define _ATL_MAX_VARTYPES 8
#endif

// Class CWebBrowser2 
// {
	void __stdcall CWebBrowser2_SetSecureLockIcon(long nSecureLockIcon)
	{
//		T* pT = static_cast<T*>(this);
//		pT->OnSetSecureLockIcon(nSecureLockIcon);
	}
	
	void __stdcall CWebBrowser2_NavigateError(IDispatch* pDisp, VARIANT* pvURL, VARIANT* pvTargetFrameName, VARIANT* pvStatusCode, VARIANT_BOOL* pbCancel)
	{
// 		T* pT = static_cast<T*>(this);
// 		ATLASSERT(V_VT(pvURL) == VT_BSTR);
// 		ATLASSERT(V_VT(pvTargetFrameName) == VT_BSTR);
// 		ATLASSERT(V_VT(pvStatusCode) == (VT_I4));
// 		ATLASSERT(pbCancel != NULL);
// 		*pbCancel=pT->OnNavigateError(pDisp,V_BSTR(pvURL),V_BSTR(pvTargetFrameName),V_I4(pvStatusCode));
	}
	
	void __stdcall CWebBrowser2_PrivacyImpactedStateChange(VARIANT_BOOL bImpacted)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnPrivacyImpactedStateChange(bImpacted);
	}
	
	void __stdcall CWebBrowser2_StatusTextChange(BSTR szText)
	{

// 		T* pT = static_cast<T*>(this);
// 		pT->OnStatusTextChange(szText);
	}
	
	void __stdcall CWebBrowser2_ProgressChange(long nProgress, long nProgressMax)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnProgressChange(nProgress, nProgressMax);
	}
	
	void __stdcall CWebBrowser2_CommandStateChange(long nCommand, VARIANT_BOOL bEnable)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnCommandStateChange(nCommand, (bEnable==VARIANT_TRUE) ? TRUE : FALSE);
	}
	
	void __stdcall CWebBrowser2_DownloadBegin(IDispatch* pDisp)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnDownloadBegin();
	}
	
	void __stdcall CWebBrowser2_DownloadComplete(IDispatch* pDisp)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnDownloadComplete();
	}
	
	void __stdcall CWebBrowser2_TitleChange(IDispatch* pDisp, BSTR szText)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnTitleChange(szText);
		SetWindowTextW( g_frmMain , (LPCWSTR)szText);
	}
	
	void __stdcall CWebBrowser2_NavigateComplete2(IDispatch* pDisp, VARIANT* pvURL)
	{

		FormMain_OnChangeUrlAndTitle();
		FormMain_OnChangeTabLabel();
// 		T* pT = static_cast<T*>(this);
// 		ATLASSERT(V_VT(pvURL) == VT_BSTR);
// 		pT->OnNavigateComplete2(pDisp, V_BSTR(pvURL));

	}
	
	void __stdcall CWebBrowser2_BeforeNavigate2(IDispatch* pDisp, VARIANT* pvURL, VARIANT* pvFlags, VARIANT* pvTargetFrameName, VARIANT* pvPostData, VARIANT* pvHeaders, VARIANT_BOOL* pbCancel)
	{


		
// 		T* pT = static_cast<T*>(this);
// 		ATLASSERT(V_VT(pvURL) == VT_BSTR);
// 		ATLASSERT(V_VT(pvTargetFrameName) == VT_BSTR);
// 		ATLASSERT(V_VT(pvPostData) == (VT_VARIANT | VT_BYREF));
// 		ATLASSERT(V_VT(pvHeaders) == VT_BSTR);
// 		ATLASSERT(pbCancel != NULL);
// 		
// 		VARIANT* vtPostedData = V_VARIANTREF(pvPostData);
// 		CSimpleArray<BYTE> pArray;
// 		if (V_VT(vtPostedData) & VT_ARRAY)
// 		{
// 			ATLASSERT(V_ARRAY(vtPostedData)->cDims == 1 && V_ARRAY(vtPostedData)->cbElements == 1);
// 			long nLowBound=0,nUpperBound=0;
// 			SafeArrayGetLBound(V_ARRAY(vtPostedData), 1, &nLowBound);
// 			SafeArrayGetUBound(V_ARRAY(vtPostedData), 1, &nUpperBound);
// 			DWORD dwSize=nUpperBound+1-nLowBound;
// 			LPVOID pData=NULL;
// 			SafeArrayAccessData(V_ARRAY(vtPostedData),&pData);
// 			ATLASSERT(pData);
// 			
// 			pArray.m_nSize=pArray.m_nAllocSize=dwSize;
// 			pArray.m_aT=(BYTE*)malloc(dwSize);
// 			ATLASSERT(pArray.m_aT);
// 			CopyMemory(pArray.GetData(), pData, dwSize);
// 			SafeArrayUnaccessData(V_ARRAY(vtPostedData));
// 		}
// 		*pbCancel=pT->OnBeforeNavigate2(pDisp, V_BSTR(pvURL), V_I4(pvFlags), V_BSTR(pvTargetFrameName), pArray, V_BSTR(pvHeaders)) ? VARIANT_TRUE : VARIANT_FALSE;
	}
	
	void __stdcall CWebBrowser2_PropertyChange(BSTR szProperty)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnPropertyChange(szProperty);
	}
	
	void __stdcall CWebBrowser2_NewWindow2(IUnknown* EventSin, IDispatch** ppDisp, VARIANT_BOOL* pbCancel)
	{

		DWORD dwThreadID = 0;
		HANDLE hThread =NULL;
		MSG				msg;
		LPDISPATCH result = NULL;
		IWebBrowser2	*pBrowserApp;
		IOleObject		*browserObject;
		WNDCLASSEX	wc	= {0};
		HRESULT hr  = -1;
		LPDISPATCH DbgppDisp = NULL;

		*pbCancel=FALSE;

		SendMessageW(
			g_frmMain,
			WM_NEW_IEVIEW,
			(WPARAM)&(msg.hwnd),
			0);

		browserObject = *((IOleObject **)GetWindowLong(msg.hwnd, GWL_USERDATA));
		hr = browserObject->lpVtbl->QueryInterface(browserObject, &IID_IWebBrowser2, (void**) &pBrowserApp);
		if ( SUCCEEDED(hr) )
		{
			hr = pBrowserApp->lpVtbl->get_Application(pBrowserApp,&DbgppDisp);
			if ( SUCCEEDED(hr) )
			{
				*ppDisp  = DbgppDisp;
			}	
		}

	}
	
	void __stdcall CWebBrowser2_DocumentComplete(IDispatch* pDisp, VARIANT* pvURL)
	{
// 		T* pT = static_cast<T*>(this);
// 		ATLASSERT(V_VT(pvURL) == VT_BSTR);
// 		pT->OnDocumentComplete(pDisp, V_BSTR(pvURL));
	}
	
	void __stdcall CWebBrowser2_OnQuit()
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnQuit();
	}
	
	void __stdcall CWebBrowser2_OnVisible(VARIANT_BOOL bVisible)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnVisible(bVisible == VARIANT_TRUE ? TRUE : FALSE);
	}
	
	void __stdcall CWebBrowser2_OnToolBar(VARIANT_BOOL bToolBar)
 	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnToolBar(bToolBar == VARIANT_TRUE ? TRUE : FALSE);
	}
	
	void __stdcall CWebBrowser2_OnMenuBar(VARIANT_BOOL bMenuBar)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnMenuBar(bMenuBar == VARIANT_TRUE ? TRUE : FALSE);
	}
	
	void __stdcall CWebBrowser2_OnStatusBar(VARIANT_BOOL bStatusBar)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnStatusBar(bStatusBar == VARIANT_TRUE ? TRUE : FALSE);
	}
	
	void __stdcall CWebBrowser2_OnFullScreen(VARIANT_BOOL bFullScreen)
	{
// 		T* pT = static_cast<T*>(this);
// 		pT->OnFullScreen(bFullScreen == VARIANT_TRUE ? TRUE : FALSE);
	}
	
	void __stdcall CWebBrowser2_OnTheaterMode(VARIANT_BOOL bTheaterMode)
	{
// 		T* pT = static_cast<T*>(this);
		// 		pT->OnTheaterMode(bTheaterMode == VARIANT_TRUE ? TRUE : FALSE);
	}
// }
	typedef struct _ATL_FUNC_INFO_
	{
		CALLCONV	cc;
		VARTYPE		vtReturn;
		SHORT		nParams;
		VARTYPE		pVarTypes[_ATL_MAX_VARTYPES];
	} _ATL_FUNC_INFO;

	typedef struct _ATL_EVENT_ENTRY_
	{
		UINT nControlID;			//ID identifying object instance
		const IID* piid;			//dispinterface IID
		int nOffset;				//offset of dispinterface from this pointer
		DISPID dispid;				//DISPID of method/property
		void *pfn;	//method to invoke
		_ATL_FUNC_INFO* pInfo;
	} _ATL_EVENT_ENTRY;


	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_StatusTextChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BSTR}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_TitleChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BSTR}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_PropertyChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BSTR}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_DownloadBeginStruct = {CC_STDCALL, VT_EMPTY, 0, {0}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_DownloadCompleteStruct = {CC_STDCALL, VT_EMPTY, 0, {0}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnQuitStruct = {CC_STDCALL, VT_EMPTY, 0, {0}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_NewWindow2Struct = {CC_STDCALL, VT_EMPTY, 2, {VT_BYREF|VT_BOOL,VT_BYREF|VT_DISPATCH}}; 
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_CommandStateChangeStruct = {CC_STDCALL, VT_EMPTY, 2, {VT_I4,VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_BeforeNavigate2Struct = {CC_STDCALL, VT_EMPTY, 7, {VT_DISPATCH,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_ProgressChangeStruct = {CC_STDCALL, VT_EMPTY, 2, {VT_I4,VT_I4}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_NavigateComplete2Struct = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, VT_BYREF|VT_VARIANT}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_DocumentComplete2Struct = {CC_STDCALL, VT_EMPTY, 2, {VT_DISPATCH, VT_BYREF|VT_VARIANT}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnVisibleStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnToolBarStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnMenuBarStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnStatusBarStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnFullScreenStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_OnTheaterModeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_SetSecureLockIconStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_I4}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_NavigateErrorStruct = {CC_STDCALL, VT_EMPTY, 5, {VT_BYREF|VT_BOOL,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_VARIANT,VT_BYREF|VT_DISPATCH}};
	__declspec(selectany) _ATL_FUNC_INFO CWebBrowser2EventsBase_PrivacyImpactedStateChangeStruct = {CC_STDCALL, VT_EMPTY, 1, {VT_BOOL}};

#define  nID 0 // ¿Ø¼þID

//Sink map is used to set up event handling
#define BEGIN_SINK_MAP(_class)\
	static const _class* _GetSinkMap()\
	{\
static const _ATL_EVENT_ENTRY map[] = {


#define SINK_ENTRY_INFO(id, iid, dispid, fn, info) {id, &iid, (int)((void*)((void*)8))-8, dispid, (void*)fn, info},
#define SINK_ENTRY_EX(id, iid, dispid, fn) SINK_ENTRY_INFO(id, iid, dispid, fn, NULL)
#define SINK_ENTRY(id, dispid, fn) SINK_ENTRY_EX(id, IID_NULL, dispid, fn)
#define END_SINK_MAP() {0, NULL, 0, 0, NULL, NULL} }; return map;}

// 	static const void* _GetSinkMap()
// 	{
// 		static const _ATL_EVENT_ENTRY map[] = 
// 		{
// 			{0, &DIID_DWebBrowserEvents2, (int)((void*)((void*)8))-8, DISPID_STATUSTEXTCHANGE, 0,&CWebBrowser2EventsBase_StatusTextChangeStruct}//,
// 			//{0, NULL, 0, 0, NULL, 0}
// 		};
// 	};

  BEGIN_SINK_MAP(void)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_STATUSTEXTCHANGE, CWebBrowser2_StatusTextChange, &CWebBrowser2EventsBase_StatusTextChangeStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_PROGRESSCHANGE, CWebBrowser2_ProgressChange, &CWebBrowser2EventsBase_ProgressChangeStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_COMMANDSTATECHANGE, CWebBrowser2_CommandStateChange, &CWebBrowser2EventsBase_CommandStateChangeStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_DOWNLOADBEGIN, CWebBrowser2_DownloadBegin, &CWebBrowser2EventsBase_DownloadBeginStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_DOWNLOADCOMPLETE, CWebBrowser2_DownloadComplete, &CWebBrowser2EventsBase_DownloadCompleteStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_TITLECHANGE, CWebBrowser2_TitleChange, &CWebBrowser2EventsBase_TitleChangeStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, CWebBrowser2_NavigateComplete2, &CWebBrowser2EventsBase_NavigateComplete2Struct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_BEFORENAVIGATE2, CWebBrowser2_BeforeNavigate2, &CWebBrowser2EventsBase_BeforeNavigate2Struct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_PROPERTYCHANGE, CWebBrowser2_PropertyChange, &CWebBrowser2EventsBase_PropertyChangeStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_NEWWINDOW2, CWebBrowser2_NewWindow2, &CWebBrowser2EventsBase_NewWindow2Struct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, CWebBrowser2_DocumentComplete, &CWebBrowser2EventsBase_DocumentComplete2Struct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONQUIT, CWebBrowser2_OnQuit, &CWebBrowser2EventsBase_OnQuitStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONVISIBLE, CWebBrowser2_OnVisible, &CWebBrowser2EventsBase_OnVisibleStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONTOOLBAR, CWebBrowser2_OnToolBar, &CWebBrowser2EventsBase_OnToolBarStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONMENUBAR, CWebBrowser2_OnMenuBar, &CWebBrowser2EventsBase_OnMenuBarStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONSTATUSBAR, CWebBrowser2_OnStatusBar, &CWebBrowser2EventsBase_OnStatusBarStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONFULLSCREEN, CWebBrowser2_OnFullScreen, &CWebBrowser2EventsBase_OnFullScreenStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_ONTHEATERMODE, CWebBrowser2_OnTheaterMode, &CWebBrowser2EventsBase_OnTheaterModeStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_SETSECURELOCKICON, CWebBrowser2_SetSecureLockIcon, &CWebBrowser2EventsBase_SetSecureLockIconStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_NAVIGATEERROR, CWebBrowser2_NavigateError, &CWebBrowser2EventsBase_NavigateErrorStruct)
	SINK_ENTRY_INFO(nID, DIID_DWebBrowserEvents2, DISPID_PRIVACYIMPACTEDSTATECHANGE, CWebBrowser2_PrivacyImpactedStateChange, &CWebBrowser2EventsBase_PrivacyImpactedStateChangeStruct)
  END_SINK_MAP()

#include "CWebevent.h"
#include "CWebBrowser2.h"

// defined (_M_IX86)
#pragma pack(push,1)
typedef	union _pfn_
{
	DWORD dwFunc;
	PVOID pfn;
}	_pfn;
typedef struct _CComStdCallThunk
{	
	void* pVtable;
	void* pFunc;
	DWORD	m_mov;          // mov dword ptr [esp+4], pThis
	DWORD   m_this;         //
	BYTE    m_jmp;          // jmp func
	DWORD   m_relproc;      // relative jmp
	//PVOID pfnCComStdCallThunk_Init;
}CComStdCallThunk, *PCComStdCallThunk;
#pragma pack(pop)

void CComStdCallThunk_Init(PCComStdCallThunk This, PVOID dw, void* pThis)
{
	_pfn pfn;
	pfn.pfn = dw;
	This->pVtable = &This->pFunc;
	This->pFunc = &This->m_mov;
	This->m_mov = 0x042444C7;
	This->m_this = (DWORD)pThis;
	This->m_jmp = 0xE9;
	This->m_relproc = (int)pfn.dwFunc - ((int)This+sizeof(CComStdCallThunk));
	FlushInstructionCache(GetCurrentProcess(), This, sizeof(CComStdCallThunk));
}

//////////////////////////////////// My IDispatch functions  //////////////////////////////////////
// The browser uses our IDispatch to give feedback when certain actions occur on the web page.
HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_AddRef(IDispatch *This)
{
	return(InterlockedIncrement(&((_IDispatchEx *)This)->refCount));
}

HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_QueryInterface(IDispatch * This, REFIID riid, void **ppvObject)
{
	*ppvObject = 0;

	if (!memcmp(riid, &IID_IUnknown, sizeof(GUID)) || !memcmp(riid, &IID_IDispatch, sizeof(GUID)))
	{
		*ppvObject = (void *)This;

		// Increment its usage count. The caller will be expected to call our
		// IDispatch's Release() (ie, Dispatch_Release) when it's done with
		// our IDispatch.
		IDispEventSimpleImpl_AddRef(This);

		return(S_OK);
	}

	*ppvObject = 0;
	return(E_NOINTERFACE);
}



HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_Release(IDispatch *This)
{
	if (InterlockedDecrement( &((_IDispatchEx *)This)->refCount ) == 0)
	{
		/* If you uncomment the following line you should get one message
		 * when the document unloads for each successful call to
		 * CreateEventHandler. If not, check you are setting all events
		 * (that you set), to null or detaching them.
		 */
		// OutputDebugString("One event handler destroyed");

		GlobalFree(((char *)This - ((_IDispatchEx *)This)->extraSize));
		return(0);
	}

	return(((_IDispatchEx *)This)->refCount);
}

HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_GetTypeInfoCount(IDispatch *This, unsigned int *pctinfo)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_GetTypeInfo(IDispatch *This, unsigned int iTInfo, LCID lcid, ITypeInfo **ppTInfo)
{
	return(E_NOTIMPL);
}

HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_GetIDsOfNames(IDispatch *This, REFIID riid, OLECHAR ** rgszNames, unsigned int cNames, LCID lcid, DISPID * rgDispId)
{
	return(S_OK);
}

static void webDetach(_IDispatchEx *lpDispatchEx)
{
// 	IHTMLWindow3	*htmlWindow3;
// 
// 	// Get the IHTMLWindow3 and call its detachEvent() to disconnect our _IDispatchEx
// 	// from the element on the web page.
// 	if (!lpDispatchEx->htmlWindow2->lpVtbl->QueryInterface(lpDispatchEx->htmlWindow2, (GUID *)&_IID_IHTMLWindow3[0], (void **)&htmlWindow3) && htmlWindow3)
// 	{
// 		htmlWindow3->lpVtbl->detachEvent(htmlWindow3, OnBeforeOnLoad, (LPDISPATCH)lpDispatchEx);
// 		htmlWindow3->lpVtbl->Release(htmlWindow3);
// 	}
// 
// 	// Release any object that was originally passed to CreateEventHandler().
// 	if (lpDispatchEx->object) lpDispatchEx->object->lpVtbl->Release(lpDispatchEx->object);
// 
// 	// We don't need the IHTMLWindow2 any more (that we got in CreateEventHandler).
// 	lpDispatchEx->htmlWindow2->lpVtbl->Release(lpDispatchEx->htmlWindow2);
}



HRESULT STDMETHODCALLTYPE Dispatch_Invoke(IDispatch *This, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, unsigned int *puArgErr)
{
	WEBPARAMS		webParams;
	BSTR			strType;
	if (251 == dispIdMember)
	{
		OutputDebugString( L"hehe" );
		
		MessageBoxW(0, L"hehe", L"hehe" ,0 );
	}
	// Get the IHTMLEventObj from the associated window.
	if (!((_IDispatchEx *)This)->htmlWindow2->lpVtbl->get_event(((_IDispatchEx *)This)->htmlWindow2, &webParams.htmlEvent) && webParams.htmlEvent)
	{
		// Get the event's type (ie, a BSTR) by calling the IHTMLEventObj's get_type().
		webParams.htmlEvent->lpVtbl->get_type(webParams.htmlEvent, &strType);
		if (strType)
		{
			// Set the WEBPARAMS.NMHDR struct's hwndFrom to the window hosting the browser,
			// and set idFrom to 0 to let him know that this is a message that was sent
			// as a result of an action occuring on the web page (that the app asked to be
			// informed of).
			webParams.nmhdr.hwndFrom = ((_IDispatchEx *)This)->hwnd;

			// Is the type "beforeunload"? If so, set WEBPARAMS.NMHDR struct's "code" to 0, otherwise 1.
			if (!(webParams.nmhdr.code = lstrcmpW(strType, &BeforeUnload[0])))
			{
				// This is a "beforeunload" event. NOTE: This should be the last event before our
				// _IDispatchEx is destroyed.

				// If the id number is negative, then this is the app's way of telling us that it
				// expects us to stuff the "eventStr" arg (passed to CreateEventHandler) into the
				// WEBPARAMS->eventStr.
				if (((_IDispatchEx *)This)->id < 0)
				{	
user:				webParams.eventStr = (LPCTSTR)((_IDispatchEx *)This)->userdata;
				}

				// We can free the BSTR we got above since we no longer need it.
				goto freestr;
			}

			// It's some other type of event. Set the WEBPARAMS->eventStr so he can do a lstrcmp()
			// to see what event happened.
			else
			{
				// Let app know that this is not a "beforeunload" event.
				webParams.nmhdr.code = 1;

				// If the app wants us to set the event type string into the WEBPARAMS, then get that
				// string as UNICODE or ANSI -- whichever is appropriate for the app window.
				if (((_IDispatchEx *)This)->id < 0) goto user;
				if (!IsWindowUnicode(webParams.nmhdr.hwndFrom))
				{
					// For ANSI, we need to convert the BSTR to an ANSI string, and then we no longer
					// need the BSTR.
					webParams.nmhdr.idFrom = WideCharToMultiByte(CP_ACP, 0, (WCHAR *)strType, -1, 0, 0, 0, 0);
					if (!(webParams.eventStr = GlobalAlloc(GMEM_FIXED, sizeof(char) * webParams.nmhdr.idFrom))) goto bad;
					WideCharToMultiByte(CP_ACP, 0, (WCHAR *)strType, -1, (char *)webParams.eventStr, webParams.nmhdr.idFrom, 0, 0);
freestr:			SysFreeString(strType);
					strType = 0;
				}

				// For UNICODE, we can use the BSTR as is. We can't free the BSTR yet.
				else
					webParams.eventStr = (LPCTSTR)strType;
			}
	
			// Send a WM_NOTIFY message to the window with the _IDispatchEx as WPARAM, and
			// the WEBPARAMS as LPARAM.
			webParams.nmhdr.idFrom = 0;
			SendMessage(webParams.nmhdr.hwndFrom, WM_NOTIFY, (WPARAM)This, (LPARAM)&webParams);

			// Free anything allocated or gotten above
bad:		if (strType) SysFreeString(strType);
			else if (webParams.eventStr && ((_IDispatchEx *)This)->id >= 0) GlobalFree((void *)webParams.eventStr);

			// If this was the "beforeunload" event, detach this IDispatch from that event for its web page element.
			// This should also cause the IE engine to call this IDispatch's Dispatch_Release().
			if (!webParams.nmhdr.code) webDetach((_IDispatchEx *)This);
		}

		// Release the IHTMLEventObj.
		webParams.htmlEvent->lpVtbl->Release(webParams.htmlEvent);
	}

	return(S_OK);
}

//Helper for finding the function index for a DISPID
HRESULT IDispEventSimpleImpl_GetFuncInfoFromId(const IID* iid, DISPID dispidMember, LCID lcid, _ATL_FUNC_INFO* info)
{
	return E_NOTIMPL;
}

HRESULT IDispEventSimpleImpl_InvokeFromFuncInfo(
												void* This,
												void *pEvent,
												_ATL_FUNC_INFO* pInfo,
												DISPPARAMS* pdispparams, VARIANT* pvarResult)
{
	CComStdCallThunk thunk;
	VARIANT tmpResult;
	HRESULT hr = 0;
	int i=0;
	VARIANTARG** pVarArgs = pInfo->nParams ? (VARIANTARG**)_alloca( sizeof(VARIANTARG*)*(pInfo->nParams) ) : 0;

	for ( i=0; i<pInfo->nParams; i++)
		pVarArgs[i] = &pdispparams->rgvarg[pInfo->nParams - i - 1];
	

	CComStdCallThunk_Init(&thunk, pEvent, This);
	
	if (pvarResult == NULL)
		pvarResult = &tmpResult;
	
	hr = DispCallFunc(
		&thunk,
		0,
		pInfo->cc,
		pInfo->vtReturn,
		pInfo->nParams,
		pInfo->pVarTypes,
		pVarArgs,
		pvarResult);

	return hr;
}
HRESULT STDMETHODCALLTYPE IDispEventSimpleImpl_Invoke(
													  IDispatch *This, 
													  DISPID dispIdMember, 
													  REFIID riid, 
													  LCID lcid, 
													  WORD wFlags, 
													  DISPPARAMS * pDispParams, 
													  VARIANT *pVarResult, 
													  EXCEPINFO *pExcepInfo, 
													  unsigned int *puArgErr)
{
//	WEBPARAMS		webParams;
//	BSTR			strType;

	const _ATL_EVENT_ENTRY* pMap = _GetSinkMap();
	const _ATL_EVENT_ENTRY* pFound = NULL;
	_ATL_FUNC_INFO info;
	_ATL_FUNC_INFO* pInfo;
	HRESULT hr = 0;

	WCHAR wszErr[255] = {0};

	wsprintf(wszErr, L"dispID = %d\n", dispIdMember);
	OutputDebugString( wszErr );
//	void (__stdcall *pEvent)() = NULL;

	while (pMap->piid != NULL)
	{
		if ((pMap->nControlID == nID) && (pMap->dispid == dispIdMember) && 
			(pMap->piid == &DIID_DWebBrowserEvents2)) //comparing pointers here should be adequate
		{
			pFound = pMap;
			break;
		}
		pMap++;
	}
	if (pFound == NULL)
		return S_OK;
	
	

	if (pFound->pInfo != NULL)
		pInfo = pFound->pInfo;
	else
	{
		pInfo = &info;
		hr = IDispEventSimpleImpl_GetFuncInfoFromId(&DIID_DWebBrowserEvents2, dispIdMember, lcid, &info);
		if (FAILED(hr))
			return S_OK;
	}
	IDispEventSimpleImpl_InvokeFromFuncInfo(This, pFound->pfn, pInfo, pDispParams, pVarResult);
		return S_OK;


	return 0;
}





DWORD COleControlSite__ConnectSink(HWND hwnd, REFIID iid, LPUNKNOWN punkSink )
{
	IOleObject		*browserObject;
	LPCONNECTIONPOINT pConnPt = NULL;
	DWORD dwCookie = 0;
	LPCONNECTIONPOINTCONTAINER pConnPtCont;
	
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));
	
	// We want to get the base address (ie, a pointer) to the IWebBrowser2 object embedded within the browser
	// object, so we can call some of the functions in the former's table.
	
	//GetEventIID( &IID_IConnectionPointContainer );

	
	if (
		(browserObject->lpVtbl != NULL) &&
 	   !((browserObject->lpVtbl)->QueryInterface( browserObject, &IID_IConnectionPointContainer,(LPVOID*)&pConnPtCont ) )
		)
	{

		if (SUCCEEDED(pConnPtCont->lpVtbl->FindConnectionPoint(pConnPtCont, iid, &pConnPt)))
		{
			pConnPt->lpVtbl->Advise(pConnPt, punkSink, &dwCookie);
			
			pConnPt->lpVtbl->Release( pConnPt );
		}
		
		pConnPtCont->lpVtbl->Release( pConnPtCont );
		
		
		return dwCookie;
	}

	return 0;

}

#define IMPLTYPE_MASK \
(IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE | IMPLTYPEFLAG_FRESTRICTED)

#define IMPLTYPE_DEFAULTSOURCE \
	(IMPLTYPEFLAG_FDEFAULT | IMPLTYPEFLAG_FSOURCE)

BOOL COleControlSite__GetEventIID(HWND hwnd, IID* piid)
{
	//*piid = GUID_NULL;
	
	
	// Use IProvideClassInfo2, if control supports it.
	IOleObject		*browserObject;
	LPPROVIDECLASSINFO2 pPCI2 = NULL;
	// Fall back on IProvideClassInfo, if IProvideClassInfo2 not supported.
	LPPROVIDECLASSINFO pPCI = NULL;
	LPTYPEINFO pClassInfo = NULL;
	LPTYPEATTR pClassAttr;		
	int nFlags;
	HREFTYPE hRefType;
	unsigned int i;
	LPTYPEATTR pEventAttr;
			
	browserObject = *((IOleObject **)GetWindowLong(hwnd, GWL_USERDATA));
	
	if (SUCCEEDED(browserObject->lpVtbl->QueryInterface(browserObject, &IID_IProvideClassInfo2,
		(LPVOID*)&pPCI2)))
	{
//		ASSERT(pPCI2 != NULL);
		
		if (SUCCEEDED(pPCI2->lpVtbl->GetGUID(pPCI2, GUIDKIND_DEFAULT_SOURCE_DISP_IID, piid)))
		{
//			ASSERT(!IsEqualIID(*piid, GUID_NULL));
		}
		else
		{
//			ASSERT(IsEqualIID(*piid, GUID_NULL));
		}
		pPCI2->lpVtbl->Release( pPCI2 );
	}
	

	
	if (//IsEqualIID(*piid, GUID_NULL) &&
		SUCCEEDED(browserObject->lpVtbl->QueryInterface(browserObject, &IID_IProvideClassInfo,
		(LPVOID*)&pPCI))
		)
	{
//		ASSERT(pPCI != NULL);

		
		if (SUCCEEDED(pPCI->lpVtbl->GetClassInfo(pPCI, &pClassInfo)))
		{
//			ASSERT(pClassInfo != NULL);
			

			if (SUCCEEDED(pClassInfo->lpVtbl->GetTypeAttr(pClassInfo, &pClassAttr)))
			{
//				ASSERT(pClassAttr != NULL);
//				ASSERT(pClassAttr->typekind == TKIND_COCLASS);
				
				// Search for typeinfo of the default events interface.

				
				for (i = 0; i < pClassAttr->cImplTypes; i++)
				{
 					if (SUCCEEDED(pClassInfo->lpVtbl->GetImplTypeFlags(pClassInfo, i, &nFlags)) &&
 						((nFlags & IMPLTYPE_MASK) == IMPLTYPE_DEFAULTSOURCE))
					{
						// Found it.  Now look at its attributes to get IID.
						
						LPTYPEINFO pEventInfo = NULL;
						
						if (SUCCEEDED(pClassInfo->lpVtbl->GetRefTypeOfImplType(
							pClassInfo,
							i,
							&hRefType)) &&
							SUCCEEDED(pClassInfo->lpVtbl->GetRefTypeInfo(
							pClassInfo,
							hRefType,
							&pEventInfo)))
						{
							//ASSERT(pEventInfo != NULL);
				
							if (SUCCEEDED(pEventInfo->lpVtbl->GetTypeAttr(pEventInfo, &pEventAttr)))
							{
								//ASSERT(pEventAttr != NULL);
								*piid = pEventAttr->guid;
								pEventInfo->lpVtbl->ReleaseTypeAttr(pEventInfo, pEventAttr);
							}
							
							pEventInfo->lpVtbl->Release(pEventInfo);
						}
						
						break;
					}
				}
				
				pClassInfo->lpVtbl->ReleaseTypeAttr(pClassInfo, pClassAttr);
			}
			
			pClassInfo->lpVtbl->Release( pClassInfo );
		}
		
		pPCI->lpVtbl->Release( pPCI );
	}

	return TRUE;
//	return (!IsEqualIID(*piid, GUID_NULL));
}









// Our IOleInPlaceFrame functions that the browser may call
HRESULT STDMETHODCALLTYPE Frame_QueryInterface(IOleInPlaceFrame FAR* This, REFIID riid, LPVOID FAR* ppvObj);
HRESULT STDMETHODCALLTYPE Frame_AddRef(IOleInPlaceFrame FAR* This);
HRESULT STDMETHODCALLTYPE Frame_Release(IOleInPlaceFrame FAR* This);
HRESULT STDMETHODCALLTYPE Frame_GetWindow(IOleInPlaceFrame FAR* This, HWND FAR* lphwnd);
HRESULT STDMETHODCALLTYPE Frame_ContextSensitiveHelp(IOleInPlaceFrame FAR* This, BOOL fEnterMode);
HRESULT STDMETHODCALLTYPE Frame_GetBorder(IOleInPlaceFrame FAR* This, LPRECT lprectBorder);
HRESULT STDMETHODCALLTYPE Frame_RequestBorderSpace(IOleInPlaceFrame FAR* This, LPCBORDERWIDTHS pborderwidths);
HRESULT STDMETHODCALLTYPE Frame_SetBorderSpace(IOleInPlaceFrame FAR* This, LPCBORDERWIDTHS pborderwidths);
HRESULT STDMETHODCALLTYPE Frame_SetActiveObject(IOleInPlaceFrame FAR* This, IOleInPlaceActiveObject *pActiveObject, LPCOLESTR pszObjName);
HRESULT STDMETHODCALLTYPE Frame_InsertMenus(IOleInPlaceFrame FAR* This, HMENU hmenuShared, LPOLEMENUGROUPWIDTHS lpMenuWidths);
HRESULT STDMETHODCALLTYPE Frame_SetMenu(IOleInPlaceFrame FAR* This, HMENU hmenuShared, HOLEMENU holemenu, HWND hwndActiveObject);
HRESULT STDMETHODCALLTYPE Frame_RemoveMenus(IOleInPlaceFrame FAR* This, HMENU hmenuShared);
HRESULT STDMETHODCALLTYPE Frame_SetStatusText(IOleInPlaceFrame FAR* This, LPCOLESTR pszStatusText);
HRESULT STDMETHODCALLTYPE Frame_EnableModeless(IOleInPlaceFrame FAR* This, BOOL fEnable);
HRESULT STDMETHODCALLTYPE Frame_TranslateAccelerator(IOleInPlaceFrame FAR* This, LPMSG lpmsg, WORD wID);

// Our IOleInPlaceFrame VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to interact with our frame window that contains
// the browser object. We must define a particular set of functions that comprise the
// IOleInPlaceFrame set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IOleInPlaceFrame structure.
IOleInPlaceFrameVtbl MyIOleInPlaceFrameTable = {Frame_QueryInterface,
Frame_AddRef,
Frame_Release,
Frame_GetWindow,
Frame_ContextSensitiveHelp,
Frame_GetBorder,
Frame_RequestBorderSpace,
Frame_SetBorderSpace,
Frame_SetActiveObject,
Frame_InsertMenus,
Frame_SetMenu,
Frame_RemoveMenus,
Frame_SetStatusText,
Frame_EnableModeless,
Frame_TranslateAccelerator};

// We need to return an IOleInPlaceFrame struct to the browser object. And one of our IOleInPlaceFrame
// functions (Frame_GetWindow) is going to need to access our window handle. So let's create our own
// struct that starts off with an IOleInPlaceFrame struct (and that's important -- the IOleInPlaceFrame
// struct *must* be first), and then has an extra data field where we can store our own window's HWND.
//
// And because we may want to create multiple windows, each hosting its own browser object (to
// display its own web page), then we need to create a IOleInPlaceFrame struct for each window. So,
// we're not going to declare our IOleInPlaceFrame struct globally. We'll allocate it later using
// GlobalAlloc, and then stuff the appropriate HWND in it then, and also stuff a pointer to
// MyIOleInPlaceFrameTable in it. But let's just define it here.
typedef struct {
	IOleInPlaceFrame	frame;		// The IOleInPlaceFrame must be first!

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IOleInPlaceFrame functions.
	// You don't want those functions to access global
	// variables, because then you couldn't use more
	// than one browser object. (ie, You couldn't have
	// multiple windows, each with its own embedded
	// browser object to display a different web page).
	//
	// So here is where I added my extra HWND that my
	// IOleInPlaceFrame function Frame_GetWindow() needs
	// to access.
	///////////////////////////////////////////////////
	HWND				window;
} _IOleInPlaceFrameEx;






// Our IOleClientSite functions that the browser may call
HRESULT STDMETHODCALLTYPE Site_QueryInterface(IOleClientSite FAR* This, REFIID riid, void ** ppvObject);
HRESULT STDMETHODCALLTYPE Site_AddRef(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE Site_Release(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE Site_SaveObject(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE Site_GetMoniker(IOleClientSite FAR* This, DWORD dwAssign, DWORD dwWhichMoniker, IMoniker ** ppmk);
HRESULT STDMETHODCALLTYPE Site_GetContainer(IOleClientSite FAR* This, LPOLECONTAINER FAR* ppContainer);
HRESULT STDMETHODCALLTYPE Site_ShowObject(IOleClientSite FAR* This);
HRESULT STDMETHODCALLTYPE Site_OnShowWindow(IOleClientSite FAR* This, BOOL fShow);
HRESULT STDMETHODCALLTYPE Site_RequestNewObjectLayout(IOleClientSite FAR* This);

// Our IOleClientSite VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to interact with our frame window that contains
// the browser object. We must define a particular set of functions that comprise the
// IOleClientSite set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IOleClientSite structure.
IOleClientSiteVtbl MyIOleClientSiteTable = {Site_QueryInterface,
Site_AddRef,
Site_Release,
Site_SaveObject,
Site_GetMoniker,
Site_GetContainer,
Site_ShowObject,
Site_OnShowWindow,
Site_RequestNewObjectLayout};






// Our IDocHostUIHandler functions that the browser may call
HRESULT STDMETHODCALLTYPE UI_QueryInterface(IDocHostUIHandler FAR* This, REFIID riid, void ** ppvObject);
HRESULT STDMETHODCALLTYPE UI_AddRef(IDocHostUIHandler FAR* This);
HRESULT STDMETHODCALLTYPE UI_Release(IDocHostUIHandler FAR* This);
HRESULT STDMETHODCALLTYPE UI_ShowContextMenu(IDocHostUIHandler FAR* This, DWORD dwID, POINT __RPC_FAR *ppt, IUnknown __RPC_FAR *pcmdtReserved, IDispatch __RPC_FAR *pdispReserved);
HRESULT STDMETHODCALLTYPE UI_GetHostInfo(IDocHostUIHandler FAR* This, DOCHOSTUIINFO __RPC_FAR *pInfo);
HRESULT STDMETHODCALLTYPE UI_ShowUI(IDocHostUIHandler FAR* This, DWORD dwID, IOleInPlaceActiveObject __RPC_FAR *pActiveObject, IOleCommandTarget __RPC_FAR *pCommandTarget, IOleInPlaceFrame __RPC_FAR *pFrame, IOleInPlaceUIWindow __RPC_FAR *pDoc);
HRESULT STDMETHODCALLTYPE UI_HideUI(IDocHostUIHandler FAR* This);
HRESULT STDMETHODCALLTYPE UI_UpdateUI(IDocHostUIHandler FAR* This);
HRESULT STDMETHODCALLTYPE UI_EnableModeless(IDocHostUIHandler FAR* This, BOOL fEnable);
HRESULT STDMETHODCALLTYPE UI_OnDocWindowActivate(IDocHostUIHandler FAR* This, BOOL fActivate);
HRESULT STDMETHODCALLTYPE UI_OnFrameWindowActivate(IDocHostUIHandler FAR* This, BOOL fActivate);
HRESULT STDMETHODCALLTYPE UI_ResizeBorder(IDocHostUIHandler FAR* This, LPCRECT prcBorder, IOleInPlaceUIWindow __RPC_FAR *pUIWindow, BOOL fRameWindow);
HRESULT STDMETHODCALLTYPE UI_TranslateAccelerator(IDocHostUIHandler FAR* This, LPMSG lpMsg, const GUID __RPC_FAR *pguidCmdGroup, DWORD nCmdID);
HRESULT STDMETHODCALLTYPE UI_GetOptionKeyPath(IDocHostUIHandler FAR* This, LPOLESTR __RPC_FAR *pchKey, DWORD dw);
HRESULT STDMETHODCALLTYPE UI_GetDropTarget(IDocHostUIHandler FAR* This, IDropTarget __RPC_FAR *pDropTarget, IDropTarget __RPC_FAR *__RPC_FAR *ppDropTarget);
HRESULT STDMETHODCALLTYPE UI_GetExternal(IDocHostUIHandler FAR* This, IDispatch __RPC_FAR *__RPC_FAR *ppDispatch);
HRESULT STDMETHODCALLTYPE UI_TranslateUrl(IDocHostUIHandler FAR* This, DWORD dwTranslate, OLECHAR __RPC_FAR *pchURLIn, OLECHAR __RPC_FAR *__RPC_FAR *ppchURLOut);
HRESULT STDMETHODCALLTYPE UI_FilterDataObject(IDocHostUIHandler FAR* This, IDataObject __RPC_FAR *pDO, IDataObject __RPC_FAR *__RPC_FAR *ppDORet);

// Our IDocHostUIHandler VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to replace/set certain user interface considerations
// (such as whether to display a pop-up context menu when the user right-clicks on the embedded
// browser object). We must define a particular set of functions that comprise the
// IDocHostUIHandler set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IDocHostUIHandler structure.
IDocHostUIHandlerVtbl MyIDocHostUIHandlerTable =  {UI_QueryInterface,
UI_AddRef,
UI_Release,
UI_ShowContextMenu,
UI_GetHostInfo,
UI_ShowUI,
UI_HideUI,
UI_UpdateUI,
UI_EnableModeless,
UI_OnDocWindowActivate,
UI_OnFrameWindowActivate,
UI_ResizeBorder,
UI_TranslateAccelerator,
UI_GetOptionKeyPath,
UI_GetDropTarget,
UI_GetExternal,
UI_TranslateUrl,
UI_FilterDataObject};

// We'll allocate our IDocHostUIHandler object dynamically with GlobalAlloc() for reasons outlined later.



// Our IOleInPlaceSite functions that the browser may call
HRESULT STDMETHODCALLTYPE InPlace_QueryInterface(IOleInPlaceSite FAR* This, REFIID riid, void ** ppvObject);
HRESULT STDMETHODCALLTYPE InPlace_AddRef(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_Release(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_GetWindow(IOleInPlaceSite FAR* This, HWND FAR* lphwnd);
HRESULT STDMETHODCALLTYPE InPlace_ContextSensitiveHelp(IOleInPlaceSite FAR* This, BOOL fEnterMode);
HRESULT STDMETHODCALLTYPE InPlace_CanInPlaceActivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceActivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_OnUIActivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_GetWindowContext(IOleInPlaceSite FAR* This, LPOLEINPLACEFRAME FAR* lplpFrame,LPOLEINPLACEUIWINDOW FAR* lplpDoc,LPRECT lprcPosRect,LPRECT lprcClipRect,LPOLEINPLACEFRAMEINFO lpFrameInfo);
HRESULT STDMETHODCALLTYPE InPlace_Scroll(IOleInPlaceSite FAR* This, SIZE scrollExtent);
HRESULT STDMETHODCALLTYPE InPlace_OnUIDeactivate(IOleInPlaceSite FAR* This, BOOL fUndoable);
HRESULT STDMETHODCALLTYPE InPlace_OnInPlaceDeactivate(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_DiscardUndoState(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_DeactivateAndUndo(IOleInPlaceSite FAR* This);
HRESULT STDMETHODCALLTYPE InPlace_OnPosRectChange(IOleInPlaceSite FAR* This, LPCRECT lprcPosRect);

// Our IOleInPlaceSite VTable. This is the array of pointers to the above functions in our C
// program that the browser may call in order to interact with our frame window that contains
// the browser object. We must define a particular set of functions that comprise the
// IOleInPlaceSite set of functions (see above), and then stuff pointers to those functions
// in their respective 'slots' in this table. We want the browser to use this VTable with our
// IOleInPlaceSite structure.
IOleInPlaceSiteVtbl MyIOleInPlaceSiteTable =  {InPlace_QueryInterface,
InPlace_AddRef,
InPlace_Release,
InPlace_GetWindow,
InPlace_ContextSensitiveHelp,
InPlace_CanInPlaceActivate,
InPlace_OnInPlaceActivate,
InPlace_OnUIActivate,
InPlace_GetWindowContext,
InPlace_Scroll,
InPlace_OnUIDeactivate,
InPlace_OnInPlaceDeactivate,
InPlace_DiscardUndoState,
InPlace_DeactivateAndUndo,
InPlace_OnPosRectChange};

// We need to pass our IOleClientSite structure to OleCreate (which in turn gives it to the browser).
// But the browser is also going to ask our IOleClientSite's QueryInterface() to return a pointer to
// our IOleInPlaceSite and/or IDocHostUIHandler structs. So we'll need to have those pointers handy.
// Plus, some of our IOleClientSite and IOleInPlaceSite functions will need to have the HWND to our
// window, and also a pointer to our IOleInPlaceFrame struct. So let's create a single struct that
// has the IOleClientSite, IOleInPlaceSite, IDocHostUIHandler, and IOleInPlaceFrame structs all inside
// it (so we can easily get a pointer to any one from any of those structs' functions). As long as the
// IOleClientSite struct is the very first thing in this custom struct, it's all ok. We can still pass
// it to OleCreate() and pretend that it's an ordinary IOleClientSite. We'll call this new struct a
// _IOleClientSiteEx.
//
// And because we may want to create multiple windows, each hosting its own browser object (to
// display its own web page), then we need to create a unique _IOleClientSiteEx struct for
// each window. So, we're not going to declare this struct globally. We'll allocate it later
// using GlobalAlloc, and then initialize the structs within it.

typedef struct {
	IOleInPlaceSite			inplace;	// My IOleInPlaceSite object. Must be first with in _IOleInPlaceSiteEx.

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IOleInPlaceSite functions.
	//
	// So here is where I added my IOleInPlaceFrame
	// struct. If you need extra variables, add them
	// at the end.
	///////////////////////////////////////////////////
	_IOleInPlaceFrameEx		frame;		// My IOleInPlaceFrame object. Must be first within my _IOleInPlaceFrameEx
} _IOleInPlaceSiteEx;

typedef struct {
	IDocHostUIHandler		ui;			// My IDocHostUIHandler object. Must be first.

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IDocHostUIHandler functions.
	///////////////////////////////////////////////////
} _IDocHostUIHandlerEx;

typedef struct {
	IOleClientSite			client;		// My IOleClientSite object. Must be first.
	_IOleInPlaceSiteEx		inplace;	// My IOleInPlaceSite object. A convenient place to put it.
	_IDocHostUIHandlerEx	ui;			// My IDocHostUIHandler object. Must be first within my _IDocHostUIHandlerEx.

	///////////////////////////////////////////////////
	// Here you add any extra variables that you need
	// to access in your IOleClientSite functions.
	///////////////////////////////////////////////////
} _IOleClientSiteEx;













// This is a simple C example. There are lots more things you can control about the browser object, but
// we don't do it in this example. _Many_ of the functions we provide below for the browser to call, will
// never actually be called by the browser in our example. Why? Because we don't do certain things
// with the browser that would require it to call those functions (even though we need to provide
// at least some stub for all of the functions).
//
// So, for these "dummy functions" that we don't expect the browser to call, we'll just stick in some
// assembly code that causes a debugger breakpoint and tells the browser object that we don't support
// the functionality. That way, if you try to do more things with the browser object, and it starts
// calling these "dummy functions", you'll know which ones you should add more meaningful code to.
#define NOTIMPLEMENTED _ASSERT(0); return(E_NOTIMPL)

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);




// A running count of how many windows we have open that contain a browser object
unsigned char WindowCount = 0;

// The class name of our Window to host the browser. It can be anything of your choosing.
static const WCHAR	ClassName[] = L"Browser Example";

// This is used by DisplayHTMLStr(). It can be global because we never change it.






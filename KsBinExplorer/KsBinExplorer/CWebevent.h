typedef struct {
	NMHDR			nmhdr;
	IHTMLEventObj *	htmlEvent;
	LPCTSTR			eventStr;
} WEBPARAMS; 

// Our _IDispatchEx struct. This is just an IDispatch with some
// extra fields appended to it for our use in storing extra
// info we need for the purpose of reacting to events that happen
// to some element on a web page.
typedef struct {
	IDispatch*		dispatchObj;	// The mandatory IDispatch.
	DWORD			refCount;		// Our reference count.
	IHTMLWindow2 *	htmlWindow2;	// Where we store the IHTMLWindow2 so that our IDispatch's Invoke() can get it.
	HWND			hwnd;			// The window hosting the browser page. Our IDispatch's Invoke() sends messages when an event of interest occurs.
	short			id;				// Any numeric value of your choosing that you wish to associate with this IDispatch.
	unsigned short	extraSize;		// Byte size of any extra fields prepended to this struct.
	IUnknown		*object;		// Some object associated with the web page element this IDispatch is for.
	void			*userdata;		// An extra pointer.
} _IDispatchEx;
static const SAFEARRAYBOUND ArrayBound = {1, 0};

static unsigned char _IID_IHTMLWindow3[] = {0xae, 0xf4, 0x50, 0x30, 0xb5, 0x98, 0xcf, 0x11, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b};
// Some misc stuff used by our IDispatch
static const BSTR	OnBeforeOnLoad = L"onbeforeunload";
static const WCHAR	BeforeUnload[] = L"beforeunload";


HRESULT STDMETHODCALLTYPE Dispatch_QueryInterface(IDispatch *, REFIID riid, void **);
HRESULT STDMETHODCALLTYPE Dispatch_AddRef(IDispatch *);
HRESULT STDMETHODCALLTYPE Dispatch_Release(IDispatch *);
HRESULT STDMETHODCALLTYPE Dispatch_GetTypeInfoCount(IDispatch *, unsigned int *);
HRESULT STDMETHODCALLTYPE Dispatch_GetTypeInfo(IDispatch *, unsigned int, LCID, ITypeInfo **);
HRESULT STDMETHODCALLTYPE Dispatch_GetIDsOfNames(IDispatch *, REFIID, OLECHAR **, unsigned int, LCID, DISPID *);
HRESULT STDMETHODCALLTYPE Dispatch_Invoke(IDispatch *, DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *, EXCEPINFO *, unsigned int *);

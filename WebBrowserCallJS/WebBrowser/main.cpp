#include "stdafx.h"
#include "resource.h"
#include "WebBrowser.h"
#include <string>

using namespace std;

INT_PTR CALLBACK DlgProc(HWND hDlg,UINT Msg,WPARAM wParam,LPARAM lParam);

HWND hMainForm;
HINSTANCE hCurrentInstance;
CWebBrowserBase *pBrowser;

void GetAppPath(wstring &sPath)
{
	sPath.resize(MAX_PATH);
	::GetModuleFileName(GetModuleHandle(NULL), (LPTSTR)sPath.c_str(), MAX_PATH);
	int index = sPath.find_last_of(L'\\');
	if(index >= 0) sPath = sPath.substr(0, index);
}

void CppCall()
{
	MessageBox(NULL, L"您调用了CppCall", L"提示(C++)", 0);
}

class ClientCall:public IDispatch
{
	long _refNum;
public:
	ClientCall()
	{
		_refNum = 1;
	}
	~ClientCall(void)
	{
	}
public:

	// IUnknown Methods

	STDMETHODIMP QueryInterface(REFIID iid,void**ppvObject)
	{
		*ppvObject = NULL;
		if (iid == IID_IUnknown)	*ppvObject = this;
		else if (iid == IID_IDispatch)	*ppvObject = (IDispatch*)this;
		if(*ppvObject)
		{
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) AddRef()
	{
		return ::InterlockedIncrement(&_refNum);
	}

	STDMETHODIMP_(ULONG) Release()
	{
		::InterlockedDecrement(&_refNum);
		if(_refNum == 0)
		{
			delete this;
		}
		return _refNum;
	}

	// IDispatch Methods

	HRESULT _stdcall GetTypeInfoCount(
		unsigned int * pctinfo) 
	{
		return E_NOTIMPL;
	}

	HRESULT _stdcall GetTypeInfo(
		unsigned int iTInfo,
		LCID lcid,
		ITypeInfo FAR* FAR* ppTInfo) 
	{
		return E_NOTIMPL;
	}

	HRESULT _stdcall GetIDsOfNames(
		REFIID riid, 
		OLECHAR FAR* FAR* rgszNames, 
		unsigned int cNames, 
		LCID lcid, 
		DISPID FAR* rgDispId 
	)
	{
		if(lstrcmp(rgszNames[0], L"CppCall")==0)
		{
			//网页调用window.external.CppCall时，会调用这个方法获取CppCall的ID
			*rgDispId = 100;
		}
		return S_OK;
	}

	HRESULT _stdcall Invoke(
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
		if(dispIdMember == 100)
		{
			//网页调用CppCall时，或根据获取到的ID调用Invoke方法
			CppCall();
		}
		return S_OK;
	}
};

typedef void _stdcall JsFunction_Callback(LPVOID pParam);

class JsFunction:public IDispatch
{
	long _refNum;
	JsFunction_Callback *m_pCallback;
	LPVOID m_pParam;
public:
	JsFunction(JsFunction_Callback *pCallback, LPVOID pParam)
	{
		_refNum = 1;
		m_pCallback = pCallback;
		m_pParam = pParam;
	}
	~JsFunction(void)
	{
	}
public:

	// IUnknown Methods

	STDMETHODIMP QueryInterface(REFIID iid,void**ppvObject)
	{
		*ppvObject = NULL;

		if (iid == IID_IUnknown)	*ppvObject = this;
		else if (iid == IID_IDispatch)	*ppvObject = (IDispatch*)this;

		if(*ppvObject)
		{
			AddRef();
			return S_OK;
		}
		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) AddRef()
	{
		return ::InterlockedIncrement(&_refNum);
	}

	STDMETHODIMP_(ULONG) Release()
	{
		::InterlockedDecrement(&_refNum);
		if(_refNum == 0)
		{
			delete this;
		}
		return _refNum;
	}

	// IDispatch Methods

	HRESULT _stdcall GetTypeInfoCount(
		unsigned int * pctinfo) 
	{
		return E_NOTIMPL;
	}

	HRESULT _stdcall GetTypeInfo(
		unsigned int iTInfo,
		LCID lcid,
		ITypeInfo FAR* FAR* ppTInfo) 
	{
		return E_NOTIMPL;
	}

	HRESULT _stdcall GetIDsOfNames(
		REFIID riid, 
		OLECHAR FAR* FAR* rgszNames, 
		unsigned int cNames, 
		LCID lcid, 
		DISPID FAR* rgDispId 
	)
	{
		//令人费解的是，网页调用函数的call方法时，没有调用GetIDsOfNames获取call的ID，而是直接调用Invoke
		return E_NOTIMPL;
	}

	HRESULT _stdcall Invoke(
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
		m_pCallback(m_pParam);
		return S_OK;
	}
};

class MainForm: public CWebBrowserBase
{
	HWND hWnd;
	ClientCall *pClientCall;
public:
	MainForm(HWND hwnd)
	{
		hWnd = hwnd;
		pClientCall = new ClientCall();
	}
	virtual HWND GetHWND(){ return hWnd;}

	static void _stdcall button1_onclick(LPVOID pParam)
	{
		MainForm *pMainForm = (MainForm*)pParam;
		MessageBox(pMainForm->hWnd, L"您点击了button1", L"提示(C++)", 0);
	}

	virtual void OnDocumentCompleted()
	{
		VARIANT params[10];

		//获取window
		IDispatch *pHtmlWindow = GetHtmlWindow();

		//获取document
		CVariant document;
		params[0].vt = VT_BSTR;
		params[0].bstrVal = L"document";
		GetProperty(pHtmlWindow, L"document", &document);

		//获取button1
		CVariant button1;
		params[0].vt = VT_BSTR;
		params[0].bstrVal = L"button1";
		InvokeMethod(document.pdispVal, L"getElementById", &button1, params, 1);

		//处理button1的onclick事件
		params[0].vt = VT_DISPATCH;
		params[0].pdispVal = new JsFunction(button1_onclick, this);
		SetProperty(button1.pdispVal, L"onclick", params);
	}

	virtual HRESULT STDMETHODCALLTYPE GetExternal(IDispatch **ppDispatch)
	{
		//重写GetExternal返回一个ClientCall对象
		*ppDispatch = pClientCall;
		return S_OK;
	}
};

int CALLBACK WinMain(HINSTANCE hInstance,HINSTANCE hPreInstance,LPSTR lpCmdLine,int ShowCmd)
{
	wstring sPath;
	GetAppPath(sPath);

	INITCOMMONCONTROLSEX icc;
	MSG msg;
	BOOL bRet;
	HACCEL hAccel;
	HICON hAppIcon;

	hCurrentInstance=hInstance;

	//初始化控件
	icc.dwSize=sizeof(icc);
	icc.dwICC=ICC_BAR_CLASSES|ICC_COOL_CLASSES;
	InitCommonControlsEx(&icc);

	//设置窗体图标
	hAppIcon=LoadIcon(hInstance,(LPCTSTR)IDI_APP);
	if(hAppIcon) SendMessage(hMainForm,WM_SETICON,ICON_BIG,(LPARAM)hAppIcon);

	//显示窗体
	hMainForm = CreateDialog(hInstance,MAKEINTRESOURCE(IDD_MAINFORM),NULL,DlgProc);
	pBrowser = new MainForm(hMainForm);
	ShowWindow(hMainForm,SW_SHOW);

	pBrowser->OpenWebBrowser();
	wstring sUrl = sPath + L"\\test.htm";
	VARIANT url;
	url.vt = VT_LPWSTR;
	url.bstrVal = (BSTR)sUrl.c_str();
	pBrowser->OpenURL(&url);

	RECT rect;
	GetClientRect(hMainForm, &rect);
	rect.top += 100;
	pBrowser->SetWebRect(&rect);

	//加载加速键
	hAccel=LoadAccelerators(hInstance,(LPCTSTR)IDR_ACCELERATOR);

	//消息循环
	while( (bRet = GetMessage( &msg, NULL, 0, 0 )) != 0)
	{ 
		if (bRet == -1)
		{
			break;
		}
		else
		{
			if (!IsDialogMessage(hMainForm, &msg) && 
				!TranslateAccelerator(hMainForm, hAccel,&msg)
				) 
			{ 
				TranslateMessage(&msg); 
				DispatchMessage(&msg); 
			}
		} 
	} 
	return 0;
}

INT_PTR CALLBACK DlgProc(HWND hDlg,UINT Msg,WPARAM wParam,LPARAM lParam)
{
	switch(Msg)
	{
	case WM_INITDIALOG:
		{
			return TRUE;
		}
	case WM_CLOSE:
		{
			PostQuitMessage(0);
			return TRUE;
		}
	case WM_COMMAND:
		{
			int wmId, wmEvent;
			wmId = LOWORD(wParam); 
			wmEvent = HIWORD(wParam); 
			switch(wmId)
			{
			case IDC_BUTTON1:
				{
					VARIANT params[10];
					VARIANT ret;
					//获取页面window
					IDispatch *pHtmlWindow = pBrowser->GetHtmlWindow();
					//页面全局函数Test实际上是window的Test方法，
					CWebBrowserBase::InvokeMethod(pHtmlWindow, L"Test", &ret, params, 0);

					break;
				}
			case IDC_BUTTON2:
				{
					VARIANT params[10];
					VARIANT ret;
					//获取页面window
					IDispatch *pHtmlWindow = pBrowser->GetHtmlWindow();
					//获取globalObject
					CVariant globalObject;
					params[0].vt = VT_BSTR;
					params[0].bstrVal = L"globalObject";
					CWebBrowserBase::GetProperty(pHtmlWindow, L"globalObject", &globalObject);
					//调用globalObject.Test
					CWebBrowserBase::InvokeMethod(globalObject.pdispVal, L"Test", &ret, params, 0);
					break;
				}
			}
			return TRUE;
		}
	case WM_SIZE:
		{
			RECT rect;
			GetClientRect(hDlg, &rect);
			rect.top += 100;
			pBrowser->SetWebRect(&rect);

			return TRUE;
		}
	}
	return 0;
}
#include "WebAutoLogin.h"

#include <MsHTML.h>
#include <atlconv.h>
#include <exdisp.h>
#include <comutil.h> // 

#pragma comment(lib, "comsupp.lib")

WebAutoLogin::WebAutoLogin(HWND hwnd, RECT webRc)
{
	m_WinContainer = std::make_shared<CAxWindow>();
	LPOLESTR pszName = OLESTR("shell.Explorer.2");
	m_WinContainer->Create(hwnd, webRc, 0, WS_CHILD | WS_VISIBLE);
	m_WinContainer->CreateControl(pszName);
	HRESULT hr = m_WinContainer->QueryControl(__uuidof(IWebBrowser2), (void**)&m_piWebBrowser);
	if (FAILED(hr))
	{
		MessageBox(hwnd, _T("获取IWebBrowser2 对象失败！！！"), _T("错误"), MB_OK | MB_ICONERROR);
		m_piWebBrowser = NULL;
	}
}

WebAutoLogin::~WebAutoLogin(void)
{
	if (NULL != m_piWebBrowser)
	{
		m_piWebBrowser->Release();
	}
}

READYSTATE WebAutoLogin::ReadyState()
{
	READYSTATE r = READYSTATE_UNINITIALIZED;
	HRESULT hr = m_piWebBrowser->get_ReadyState(&r);
	//printf("get_ReadyState = %d",r);
	if (SUCCEEDED(hr) && r == READYSTATE_COMPLETE)
	{

	}
	return r;
}

bool WebAutoLogin::AutoLogin(HWND hwnd, std::wstring userName, std::wstring password)
{
	bool isLogin = false;
	HRESULT hr = S_OK;
	IHTMLElement *user_nameElet = GetHTMLElementByIdOrName(L"username");
	if (user_nameElet != 0)
	{
		//转换成CComBSTR      
		CComBSTR bStr = userName.c_str();
		//输入内容     
		hr = user_nameElet->put_innerText(bStr);
		user_nameElet->Release();
	}
	else
	{
		MessageBox(hwnd, _T("获取:用户名HTMLElement 失败!"), _T("错误"), MB_OK | MB_ICONERROR);
		return isLogin;
	}

	IHTMLElement *passwdElet = GetHTMLElementByIdOrName(L"password");
	if (passwdElet != 0)
	{
		//转换成CComBSTR      
		CComBSTR bStr = password.c_str();
		//输入内容     
		hr = passwdElet->put_innerText(bStr);
		passwdElet->Release();
	}
	else
	{
		MessageBox(hwnd, _T("获取:密码HTMLElement 失败!"), _T("错误"), MB_OK | MB_ICONERROR);
		return isLogin;
	}

	IHTMLElement *loginSubElet = GetHTMLElementByTag(L"input", L"value", L"登 录");
	if (loginSubElet != 0)
	{
		loginSubElet->click();
		loginSubElet->Release();
		isLogin = true;
	}
	else
	{
		MessageBox(hwnd, _T("获取:登陆HTMLElement 失败!"), _T("错误"), MB_OK | MB_ICONERROR);
		return isLogin;
	}
	return isLogin;
}

bool WebAutoLogin::BaiduSearch(HWND hwnd, std::wstring keyWord)
{
	bool isLogin = false;
	HRESULT hr = S_OK;
	// 输入框
	IHTMLElement *user_nameElet = GetHTMLElementByIdOrName(L"wd");
	if (user_nameElet != 0)
	{
		//转换成CComBSTR      
		CComBSTR bStr = keyWord.c_str();
		//输入内容     
		hr = user_nameElet->put_innerText(bStr);
		user_nameElet->Release();
	}
	else
	{
		MessageBox(hwnd, _T("获取:关键字 失败!"), _T("错误"), MB_OK | MB_ICONERROR);
	}

	// 搜索按钮
	IHTMLElement *loginSubElet = GetHTMLElementByIdOrName(L"su");
	if (loginSubElet != 0)
	{
		loginSubElet->click();
		loginSubElet->Release();
		isLogin = true;
	}
	else
	{
		MessageBox(hwnd, _T("获取:搜索按钮 失败!"), _T("错误"), MB_OK | MB_ICONERROR);
	}
	return isLogin;
}

bool WebAutoLogin::LoginResult()
{
	bool isLogin = false;
	IDispatch *dispatch = 0;
	HRESULT hr = m_piWebBrowser->get_Document(&dispatch);
	if ((S_OK == hr) && (dispatch != 0))
	{
		IHTMLDocument2 *doc;
		hr = dispatch->QueryInterface(IID_IHTMLDocument2, (void**)&doc);
		dispatch->Release();
		if (S_OK == hr)
		{
			//登陆成功的 判断方式可以用不同的 方法
			BSTR bstrCookie;
			hr = doc->get_cookie(&bstrCookie);
			if (S_OK == hr)
			{
				_bstr_t bstr_t(bstrCookie);
				std::string strCookie(bstr_t);
				::SysFreeString(bstrCookie);
			}
			BSTR bstrReferrer;
			hr = doc->get_referrer(&bstrReferrer);
			if (S_OK == hr)
			{
				if (NULL != bstrReferrer)
				{
					_bstr_t bstr_t0(bstrReferrer);
					std::string strReferrer(bstr_t0);
					::SysFreeString(bstrReferrer);
				}
			}
			BSTR bstrUrl;
			hr = doc->get_URL(&bstrUrl);
			if (S_OK == hr && NULL != bstrUrl)
			{
				_bstr_t bstr_t(bstrUrl);
				std::string strUrl(bstr_t);
				if (0 == strcmp("http://www.csdn.net/", strUrl.c_str()))
				{
					isLogin = true;
				}
				// free the BSTR
				::SysFreeString(bstrUrl);
			}
		}
		doc->Release();
	}
	dispatch->Release();

	return isLogin;
}

void WebAutoLogin::Navigate(std::wstring strUrl)
{
	VARIANT varMyURL;
	VariantInit(&varMyURL);
	varMyURL.vt = VT_BSTR;
	varMyURL.bstrVal = SysAllocString(strUrl.c_str());
	m_piWebBrowser->Navigate2(&varMyURL, 0, 0, 0, 0);
	SysFreeString(varMyURL.bstrVal);
	VariantClear(&varMyURL);
}

IHTMLElement * WebAutoLogin::GetHTMLElementByTag(std::wstring tagName, std::wstring PropertyName,
	std::wstring macthValue)
{
	IHTMLElement *retElement = 0;
	IDispatch *dispatch = 0;
	HRESULT hr = m_piWebBrowser->get_Document(&dispatch);
	if ((S_OK == hr) && (0 != dispatch))
	{
		IHTMLDocument2 *doc;
		dispatch->QueryInterface(IID_IHTMLDocument2, (void**)&doc);
		dispatch->Release();
		IHTMLElementCollection* doc_all;
		hr = doc->get_all(&doc_all);      // this is like doing document.all
		if (S_OK == hr)
		{
			VARIANT vKey;
			vKey.vt = VT_BSTR;
			vKey.bstrVal = SysAllocString(tagName.c_str());
			VARIANT vIndex;
			VariantInit(&vIndex);
			hr = doc_all->tags(vKey, &dispatch);       // this is like doing document.all["messages"]
			//清理
			SysFreeString(vKey.bstrVal);
			VariantClear(&vKey);
			VariantClear(&vIndex);
			if ((S_OK == hr) && (0 != dispatch))
			{
				CComQIPtr< IHTMLElementCollection > all_tags = dispatch;
				//hr = dispatch->QueryInterface(IHTMLElementCollection,(void **)&all_tags); // it's the caller's responsibility to release 
				if (S_OK == hr)
				{
					long nTagsCount = 0; //
					hr = all_tags->get_length(&nTagsCount);
					if (FAILED(hr))
					{
						return retElement;
					}

					for (long i = 0; i < nTagsCount; i++)
					{
						CComDispatchDriver spInputElement; //取得第 i 项
						hr = all_tags->item(CComVariant(i), CComVariant(i), &spInputElement);

						if (FAILED(hr))
							continue;
						CComVariant vValue;
						hr = spInputElement.GetPropertyByName(PropertyName.c_str(), &vValue);
						if (VT_EMPTY != vValue.vt)
						{
							LPCTSTR lpValue = vValue.bstrVal ?
								OLE2CT(vValue.bstrVal) : NULL;
							if (NULL == lpValue)
								continue;
							std::wstring cs = (LPCTSTR)lpValue;
							if (0 == _tcscmp(cs.c_str(), macthValue.c_str()))
							{
								hr = spInputElement->QueryInterface(IID_IHTMLElement, (void **)&retElement);
								if (S_OK == hr)
								{
								}
								else
								{
									retElement = 0;
								}
								break;
							}
						}
						//
						//CComVariant vName,vVal,vType; //名，值，类型
						//hr = spInputElement.GetPropertyByName( L"name", &vName );
						//if( FAILED( hr ) ) continue;
						//hr = spInputElement.GetPropertyByName( L"value", &vVal );
						//if( FAILED( hr ) ) continue;
						//hr = spInputElement.GetPropertyByName( L"type", &vType );
						//if( FAILED( hr ) ) continue;
						//LPCTSTR lpName = vName.bstrVal?
						//	OLE2CT( vName.bstrVal ) : _T("NULL"); //未知域名
						//LPCTSTR lpVal  = vVal.bstrVal?
						//	OLE2CT( vVal.bstrVal  ) : _T("NULL"); //空值，未输入
						//LPCTSTR lpType = vType.bstrVal?
						//	OLE2CT( vType.bstrVal ) : _T("NULL"); //未知类型
					}
				}
				else
				{
					retElement = 0;
				}
				dispatch->Release();
			}
			doc_all->Release();
		}
		doc->Release();
	}
	return retElement;
}

IHTMLElement * WebAutoLogin::GetHTMLElementByIdOrName(std::wstring idorName)
{
	IHTMLElement *retElement = 0;
	IDispatch *dispatch = 0;
	HRESULT hr = m_piWebBrowser->get_Document(&dispatch);
	if ((S_OK == hr) && (0 != dispatch))
	{
		IHTMLDocument2 *doc;
		dispatch->QueryInterface(IID_IHTMLDocument2, (void**)&doc);
		dispatch->Release();
		IHTMLElementCollection* doc_all;
		hr = doc->get_all(&doc_all);      // this is like doing document.all
		if (S_OK == hr)
		{
			VARIANT vKey;
			vKey.vt = VT_BSTR;
			vKey.bstrVal = SysAllocString(idorName.c_str());
			VARIANT vIndex;
			VariantInit(&vIndex);
			hr = doc_all->item(vKey, vIndex, &dispatch);       // this is like doing document.all["messages"]
			//清理
			SysFreeString(vKey.bstrVal);
			VariantClear(&vKey);
			VariantClear(&vIndex);
			if ((S_OK == hr) && (0 != dispatch))
			{
				hr = dispatch->QueryInterface(IID_IHTMLElement, (void **)&retElement); // it's the caller's responsibility to release 
				if (S_OK == hr)
				{
				}
				else
				{
					retElement = 0;
				}
				dispatch->Release();
			}
			doc_all->Release();
		}
		doc->Release();
	}
	return retElement;
}

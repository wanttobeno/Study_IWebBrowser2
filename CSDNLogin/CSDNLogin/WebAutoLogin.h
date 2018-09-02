#pragma once

#include <windows.h>
#include <MsHTML.h>
#include <string>
#include <ExDisp.h>

#include <memory>

#include <atlbase.h>
#include <atlwin.h>

class WebAutoLogin
{
public:
	explicit WebAutoLogin(HWND hwnd, RECT webRc);

	~WebAutoLogin();

	READYSTATE ReadyState();

	bool AutoLogin(HWND hwnd, std::wstring userName, std::wstring password);

	bool BaiduSearch(HWND hwnd, std::wstring keyWord);

	bool LoginResult();

	void Navigate(std::wstring strUrl);

	IHTMLElement * GetHTMLElementByTag(std::wstring tagName, std::wstring PropertyName,
		std::wstring macthValue);

	IHTMLElement * GetHTMLElementByIdOrName(std::wstring idorName);
private:

	std::shared_ptr<IWebBrowser2> m_piWebBrowser; //ä¯ÀÀÆ÷Ö¸Õë
	std::shared_ptr<CAxWindow> m_WinContainer;
};


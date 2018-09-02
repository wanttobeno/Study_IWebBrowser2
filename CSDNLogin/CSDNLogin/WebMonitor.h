#pragma once
#include <ExDisp.h>
#include <MsHTML.h>
#include <string>
#include <vector>

class HtmlElementSink;

class TDocHostUIHandlerImpl;

class WebMonitor
{
public:
	WebMonitor();
	~WebMonitor();


	void ConnectHTMLElementEvent(IHTMLElement* pElem);


	bool ExecJsFun(const std::wstring& lpJsFun, const std::vector<std::wstring>& params);


	void RegisterUIHandlerToJs();
protected:
	IWebBrowser2*		m_piWebBrowser;
	HtmlElementSink*	m_pHtmlElementSink;
	TDocHostUIHandlerImpl* m_pDocHostUIHandler;

};


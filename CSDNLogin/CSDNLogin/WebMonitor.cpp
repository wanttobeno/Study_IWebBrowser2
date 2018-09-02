#include "WebMonitor.h"
#include "HtmlElementSink.h"

#include "TDocHostUIHandlerImpl.h"

#include <ExDisp.h>
#include <atlcomcli.h>
#include <MsHtmHst.h>

WebMonitor::WebMonitor()
{
}

WebMonitor::~WebMonitor()
{
}

//监听  页面事件 
void WebMonitor::ConnectHTMLElementEvent(IHTMLElement* pElem)
{
	if (NULL == pElem)
	{
		return;
	}
	HRESULT hr;
	IConnectionPointContainer* pCPC = NULL;
	IConnectionPoint* pCP = NULL;
	DWORD dwCookie;

	// Check that this is a connectable object.
	hr = pElem->QueryInterface(IID_IConnectionPointContainer, (void**)&pCPC);

	if (SUCCEEDED(hr))
	{
		// 枚举 查看一下 支持的 ConnectionPoint
		//IEnumConnectionPoints *pECPS = NULL;
		//hr = pCPC->EnumConnectionPoints(&pECPS);
		//if (SUCCEEDED(hr))
		//{
		//	IConnectionPoint* tmppCP = NULL;
		//	ULONG cFetched = 1;
		//	IID guid;
		//	while(SUCCEEDED(hr = pECPS->Next(1,&tmppCP,&cFetched))
		//		&&(cFetched>0))
		//	{	
		//		tmppCP->GetConnectionInterface(&guid);
		//	}
		//}

		// Find the connection point.
		//hr = pCPC->FindConnectionPoint(DIID_HTMLElementEvents2, &pCP);
		hr = pCPC->FindConnectionPoint(DIID_HTMLButtonElementEvents, &pCP);

		if (SUCCEEDED(hr))
		{
			// Advise the connection point.
			// pUnk is the IUnknown interface pointer for your event sink
			if (NULL == m_pHtmlElementSink)
			{
				m_pHtmlElementSink = new HtmlElementSink(this);
			}
			hr = pCP->Advise(m_pHtmlElementSink, &dwCookie);

			if (SUCCEEDED(hr))
			{
				// Successfully advised
			}
			pCP->Release();
		}
		pCPC->Release();
	}
}

bool WebMonitor::ExecJsFun(const std::wstring& lpJsFun, const std::vector<std::wstring>& params)
{
	if (NULL == m_piWebBrowser)
		return false;
	CComPtr<IDispatch> pDoc;
	HRESULT hr = m_piWebBrowser->get_Document(&pDoc);
	if (FAILED(hr))
		return false;
	CComQIPtr<IHTMLDocument2> pDoc2 = pDoc;
	if (NULL == pDoc2)
		return false;
	CComQIPtr<IDispatch> pScript;
	hr = pDoc2->get_Script(&pScript);
	if (FAILED(hr))
		return false;
	DISPID id = NULL;
	CComBSTR bstrFun(lpJsFun.c_str());
	hr = pScript->GetIDsOfNames(IID_NULL, &bstrFun, 1, LOCALE_SYSTEM_DEFAULT, &id);
	if (FAILED(hr))
		return false;
	DISPPARAMS dispParams;
	memset(&dispParams, 0, sizeof(DISPPARAMS));
	int nParamCount = params.size();
	if (nParamCount > 0)
	{
		dispParams.cArgs = nParamCount;
		dispParams.rgvarg = new VARIANT[nParamCount];
		for (int i = 0; i < nParamCount; ++i)
		{
			const std::wstring& str = params[nParamCount - 1 - i];
			CComBSTR bstr(str.c_str());
			bstr.CopyTo(&dispParams.rgvarg[i].bstrVal);
			dispParams.rgvarg[i].vt = VT_BSTR;
		}
	}

	EXCEPINFO execInfo;
	memset(&execInfo, 0, sizeof(EXCEPINFO));
	VARIANT vResult;
	UINT uArgError = (UINT)-1;
	hr = pScript->Invoke(id, IID_NULL, 0, DISPATCH_METHOD, &dispParams, &vResult, &execInfo, &uArgError);
	delete[] dispParams.rgvarg;
	if (FAILED(hr))
		return false;
	return true;
}

void WebMonitor::RegisterUIHandlerToJs()
{
	ICustomDoc   *m_spCustDoc;
	HRESULT   hr;
	CComPtr<IDispatch> pDoc;
	hr = m_piWebBrowser->get_Document(&pDoc);
	if (FAILED(hr))
		return;
	CComQIPtr<IHTMLDocument2> pDoc2 = pDoc;
	if (NULL == pDoc2)
		return;
	hr = pDoc2->QueryInterface(IID_ICustomDoc, (void**)&m_spCustDoc);
	if (SUCCEEDED(hr))
	{
		if (NULL == m_pDocHostUIHandler)
		{
			m_pDocHostUIHandler = new TDocHostUIHandlerImpl();
		}
		hr = m_spCustDoc->SetUIHandler(m_pDocHostUIHandler);
		if (SUCCEEDED(hr))
		{
			// Successfully advised
		}
	}
}

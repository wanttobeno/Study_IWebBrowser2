#include <Windows.h>
#include <tchar.h>

#include "WebBrowserCore.h"


HINSTANCE handleInstance;
HWND mainWindowHandle;
TCHAR* mainWindowTitle = _T("Embedded WebBrowser");
TCHAR* mainWindowClassName = _T("WndClsEmbeddedWebBrowser");

ATOM RegMainClass();
LRESULT CALLBACK mainWindowProcess(HWND hWnd,
	UINT uMsg, WPARAM wParam, LPARAM lParam);


WebBrowser *webBrowserPtr = nullptr;

HRESULT testFunction(DISPID id, DISPPARAMS* args, VARIANT* s) {
	MessageBox(NULL, _T("javascript call without argument"),
		_T("call"),
		MB_ICONERROR);
	return S_OK;
}
HRESULT testFunction1Arg(DISPID id, DISPPARAMS* args, VARIANT* s) {
	MessageBox(NULL, _T("javascript call with one argument"),
		_T("call"),
		MB_ICONERROR);
	return S_OK;
}


INT WINAPI WinMain(HINSTANCE hInstance,
				   HINSTANCE hPrevInstance,
				   LPSTR lpCmdLine,
				   INT nCmdShow)
{

	OleInitialize(NULL);

	MSG msg;

	handleInstance = hInstance;

	if (!RegMainClass())
	{
		MessageBox(NULL, _T("RegMainClass failed"),
			         _T("WinMain"),
					 MB_ICONERROR);
		return 1;
	}

	mainWindowHandle = CreateWindowEx(0, mainWindowClassName,
						      mainWindowTitle,
							  WS_OVERLAPPEDWINDOW,
							  CW_USEDEFAULT, CW_USEDEFAULT,
						      CW_USEDEFAULT, CW_USEDEFAULT,
							  NULL, NULL, handleInstance, NULL);

	ShowWindow(mainWindowHandle, nCmdShow);

	RECT mainClient;
	GetClientRect(mainWindowHandle, &mainClient);

	webBrowserPtr = new WebBrowser(mainWindowHandle);
	if (webBrowserPtr != nullptr) {
		RECT webClient;
		webClient.left = 0;
		webClient.top = 0;
		webClient.right = mainClient.right;
		webClient.bottom = mainClient.bottom;
		webBrowserPtr->SetRect(webClient);

		webBrowserPtr->AddJavaScriptFunction(L"testFunction", &testFunction);
		webBrowserPtr->AddJavaScriptFunction(L"testFunction1Arg", &testFunction1Arg);
	}

	
	
	while (GetMessage(&msg, NULL, 0, 0))
	{
		//enable tab delete and other keys on web page
		webBrowserPtr->TranslateMessage(&msg);

		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	return 0;
}

ATOM RegMainClass()
{
	WNDCLASS wc;
	wc.cbClsExtra = wc.cbWndExtra = 0;
	wc.hbrBackground = (HBRUSH) 5;
	wc.hCursor = LoadCursor(NULL, IDC_ARROW);
	wc.hIcon = NULL;
	wc.hInstance = handleInstance;
	wc.lpfnWndProc = mainWindowProcess;
	wc.lpszClassName = mainWindowClassName;
	wc.lpszMenuName = NULL;
	wc.style = 0;

	return RegisterClass(&wc);
}



LRESULT CALLBACK mainWindowProcess(HWND hWnd, UINT uMsg,
	WPARAM wParam, LPARAM lParam)
{

	switch (uMsg)
	{
	
	case WM_SIZE:
		if (webBrowserPtr != nullptr)
		{
			RECT mainClient;
			GetClientRect(mainWindowHandle, &mainClient);

			RECT webClient;
			webClient.left = 0;
			webClient.top = 0;
			webClient.right = mainClient.right;
			webClient.bottom = mainClient.bottom;
			if (webBrowserPtr != nullptr)
				webBrowserPtr->SetRect(webClient);
		}
		break;
	case WM_DESTROY:
		ExitProcess(0);
		break;
	default:
		return DefWindowProc(hWnd, uMsg, wParam, lParam);
	}
}
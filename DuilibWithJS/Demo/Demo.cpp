// Demo.cpp : 定义应用程序的入口点。
//

#include "stdafx.h"
#include "Demo.h"
#include "MainFrameWnd.h"

int APIENTRY _tWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPTSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
	CoInitialize(NULL);
	OleInitialize(NULL);
	// 初始化UI管理器
	CPaintManagerUI::SetInstance(hInstance);
	// 资源路径
	CDuiString strResourcePath = CPaintManagerUI::GetInstancePath();
	strResourcePath += _T("Res\\");
	CPaintManagerUI::SetResourcePath(strResourcePath.GetData());
	// 主窗口
	CMainFrameWnd* pFrame = new CMainFrameWnd();
	pFrame->Create(NULL, _T("MainFrameWnd"), UI_WNDSTYLE_FRAME, WS_EX_STATICEDGE | WS_EX_APPWINDOW);
	::ShowWindow(*pFrame, SW_SHOW);
	CPaintManagerUI::MessageLoop();

	CResourceManager::GetInstance()->Release();

	OleUninitialize();
	CoUninitialize();

	return 0;
}
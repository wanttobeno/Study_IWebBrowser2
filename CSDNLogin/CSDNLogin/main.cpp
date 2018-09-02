#include <windows.h>
#include "resource.h"
#include "WebAutoLogin.h"
#include <tchar.h>
#include <MsHTML.h>
#include <ExDisp.h>


std::shared_ptr<WebAutoLogin> gWebAutoLogin;

bool bIsSimulateLogin = false;

CComModule _Module;
extern __declspec(selectany) CAtlModule* _pAtlModule = &_Module;

void StartWaitWebLoad(HWND hwnd)
{

}

LRESULT DialogInit(HWND hDlg, WPARAM wParam, LPARAM lParam)
{
	RECT rc;
	GetClientRect(hDlg, &rc);

	HWND hStaticOption = GetDlgItem(hDlg, IDC_STATIC_OPTION);
	RECT rc_Option;
	GetWindowRect(hStaticOption, &rc_Option);
	//
	SetDlgItemText(hDlg, IDC_EDIT_URL, _T("https://passport.csdn.net/account/login"));
	SetDlgItemText(hDlg, IDC_EDIT_NAME, _T("xxxxxxx"));
	SetDlgItemText(hDlg, IDC_EDIT_PWD, _T("xxxxxx"));
	//
	RECT webRc;
	webRc.left = 0;
	webRc.top = rc_Option.bottom;
	webRc.right = rc.right;
	webRc.bottom = rc.bottom;
	gWebAutoLogin = std::make_shared<WebAutoLogin>(hDlg, webRc);
	gWebAutoLogin->Navigate(L"https://www.baidu.com/");
	StartWaitWebLoad(hDlg);
	return TRUE;
}

void LoadWeb(HWND hDlg)
{
	wchar_t url[MAX_PATH * 2 + 1] = { 0 };
	GetDlgItemText(hDlg, IDC_EDIT_URL, url, MAX_PATH * 2);
	std::wstring strUrl;
	strUrl.assign(url);

	gWebAutoLogin->Navigate(strUrl);
	StartWaitWebLoad(hDlg);
}

void SimulateLogin(HWND hDlg)
{
	wchar_t name[MAX_PATH + 1] = { 0 };
	GetDlgItemText(hDlg, IDC_EDIT_NAME, name, MAX_PATH);
	wchar_t pswd[MAX_PATH + 1] = { 0 };
	GetDlgItemText(hDlg, IDC_EDIT_PWD, pswd, MAX_PATH);

	std::wstring userName;
	userName.assign(name);
	std::wstring password;
	password.assign(pswd);

	if (gWebAutoLogin->AutoLogin(hDlg, userName, password))
	{
		bIsSimulateLogin = TRUE;
		StartWaitWebLoad(hDlg);
	}
}

void SimulateBaidu(HWND hDlg)
{
	wchar_t name[MAX_PATH + 1] = { 0 };
	GetDlgItemText(hDlg, IDC_EDIT_NAME, name, MAX_PATH);
	std::wstring keyWord;
	keyWord.assign(name);
	if (gWebAutoLogin->BaiduSearch(hDlg, keyWord))
	{
		StartWaitWebLoad(hDlg);
	}
}

LRESULT CommandFun(HWND hDlg, WPARAM wParam, LPARAM lParam)
{
	if (LOWORD(wParam) == IDOK)
	{
		PostQuitMessage(0);
	}
	else if (LOWORD(wParam) == IDC_BTN_GO)
	{
		LoadWeb(hDlg);
	}
	else if (LOWORD(wParam) == IDC_BTN_LOGIN)
	{
		SimulateLogin(hDlg);
	}
	else if (LOWORD(wParam) == IDC_BTN_BAIDU)
	{
		SimulateBaidu(hDlg);
	}
	return TRUE;
}


//   Windows   事件处理
LRESULT CALLBACK MainDialogProc(HWND hDlg, UINT message,
	WPARAM wParam, LPARAM lParam)
{
	//消息的处理，我想你要的就是这里了
	switch (message)
	{
	case WM_INITDIALOG:
		return DialogInit(hDlg, wParam, lParam);
		break;
	case WM_COMMAND:
		return CommandFun(hDlg, wParam, lParam);
		break;
	case WM_CLOSE://关闭在这里
		EndDialog(hDlg, TRUE);
		return TRUE;
		break;
	}
	return FALSE;
}



int WINAPI _tWinMain(HINSTANCE hInst, HINSTANCE h0, LPTSTR lpCmdLine, int nCmdShow)
{
	CoInitialize(NULL);
	OleInitialize(NULL);
	DialogBox(hInst, MAKEINTRESOURCE(IDD_DIALOG_MAIN), NULL, (DLGPROC)MainDialogProc);
	return TRUE;
}
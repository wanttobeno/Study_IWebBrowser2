#include <Windows.h>
#include <tchar.h>

#include "WebBrowser.h"

#define PURE_POPUP_WINDOW

HINSTANCE hInst;
HWND hWndMain;
TCHAR* szWndTitleMain = _T("Embedded WebBrowser in Pure C++");
TCHAR* szWndClassMain = _T("WndClass_EmbeddedWB_PureCPP");

ATOM RegMainClass();
LRESULT CALLBACK WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

HWND hWndAddressBar;

#define btnBack 1
#define btnNext 2
#define btnRefresh 3
#define btnGo 4

WebBrowser *g_pWebBrowser = nullptr;
LPCWSTR g_lpszCmdLine = nullptr;

enum
{
    WM_DO_NAVIGATE = WM_USER + 100,
};

INT WINAPI _tWinMain(HINSTANCE hInstance,
    HINSTANCE hPrevInstance,
    LPTSTR lpCmdLine,
    INT nCmdShow)
{
    g_lpszCmdLine = lpCmdLine;

    OleInitialize(NULL);

    MSG msg;

    hInst = hInstance;

    if (!RegMainClass())
    {
        MessageBox(NULL, _T("Cannot register main window class"),
            _T("Error No. 1"),
            MB_ICONERROR);
        return 1;
    }

    hWndMain = CreateWindowEx(
#ifdef PURE_POPUP_WINDOW
        WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
#else
        0,
#endif
        szWndClassMain,
        szWndTitleMain,
#ifdef PURE_POPUP_WINDOW
        WS_POPUP, 400, 300, 800, 600,
#else
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        CW_USEDEFAULT, CW_USEDEFAULT,
#endif
        NULL, NULL, hInst, NULL);

    RECT rcClient;
    GetClientRect(hWndMain, &rcClient);

    g_pWebBrowser = new WebBrowser(hWndMain);
    RECT rc;
    rc.left = 0;
#ifndef PURE_POPUP_WINDOW
    rc.top = 45;
#else
    rc.top = 0;
#endif
    rc.right = rcClient.right;
    rc.bottom = rcClient.bottom;
    g_pWebBrowser->SetRect(rc);

    ShowWindow(hWndMain, nCmdShow);

    while (GetMessage(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}

ATOM RegMainClass()
{
    WNDCLASS wc;
    wc.cbClsExtra = wc.cbWndExtra = 0;
    wc.hbrBackground = (HBRUSH)COLOR_WINDOW;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hIcon = NULL;
    wc.hInstance = hInst;
    wc.lpfnWndProc = WndProc;
    wc.lpszClassName = szWndClassMain;
    wc.lpszMenuName = NULL;
    wc.style = 0;

    return RegisterClass(&wc);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_CREATE:
#ifndef PURE_POPUP_WINDOW
        CreateWindowEx(0, _T("BUTTON"),
            _T("<<< Back"),
            WS_CHILD | WS_VISIBLE,
            5, 5,
            80, 30,
            hWnd, (HMENU)btnBack, hInst, NULL);

        CreateWindowEx(0, _T("BUTTON"),
            _T("Next >>>"),
            WS_CHILD | WS_VISIBLE,
            90, 5,
            80, 30,
            hWnd, (HMENU)btnNext, hInst, NULL);

        CreateWindowEx(0, _T("BUTTON"),
            _T("Refresh"),
            WS_CHILD | WS_VISIBLE,
            175, 5,
            80, 30,
            hWnd, (HMENU)btnRefresh, hInst, NULL);

        hWndAddressBar =
            CreateWindowEx(0, _T("EDIT"),
                _T("http://news.sogou.com/"),
                WS_CHILD | WS_VISIBLE | WS_BORDER,
                260, 10,
                200, 20,
                hWnd, NULL, hInst, NULL);

        CreateWindowEx(0, _T("BUTTON"),
            _T("Go"),
            WS_CHILD | WS_VISIBLE,
            465, 5,
            80, 30,
            hWnd, (HMENU)btnGo, hInst, NULL);
#else
        PostMessage(hWnd, WM_DO_NAVIGATE, 0, 0);
#endif
        break;
    case WM_COMMAND:
#ifndef PURE_POPUP_WINDOW
        switch (LOWORD(wParam))
        {
        case btnBack:
            g_pWebBrowser->GoBack();
            break;
        case btnNext:
            g_pWebBrowser->GoForward();
            break;
        case btnRefresh:
            g_pWebBrowser->Refresh();
            break;
        case btnGo:
            TCHAR * szUrl = new TCHAR[1024];
            GetWindowText(hWndAddressBar, szUrl, 1024);
            g_pWebBrowser->Navigate(szUrl);
            break;
        }
#endif
        break;
    case WM_SIZE:
        if (g_pWebBrowser != 0)
        {
            RECT rcClient;
            GetClientRect(hWndMain, &rcClient);

            RECT rc;
            rc.left = 0;
#ifndef PURE_POPUP_WINDOW
            rc.top = 45;
#else
            rc.top = 0;
#endif
            rc.right = rcClient.right;
            rc.bottom = rcClient.bottom;
            if (g_pWebBrowser != 0)
                g_pWebBrowser->SetRect(rc);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    case WM_DO_NAVIGATE:
        if (g_pWebBrowser)
        {
            if (g_lpszCmdLine && g_lpszCmdLine[0])
                g_pWebBrowser->Navigate(g_lpszCmdLine);
            else
                g_pWebBrowser->Navigate(L"https://www.baidu.com/");
        }
        break;
    default:
        return DefWindowProc(hWnd, uMsg, wParam, lParam);
    }

    return 0;
}

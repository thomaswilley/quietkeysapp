// QuietkeysApp.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "QuietkeysApp.h"
#include <shellapi.h>
#include <stdio.h>
#include <strsafe.h>

#define MAX_LOADSTRING 100
#define APPWM_ICONNOTIFY (WM_APP + 1)

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

namespace gQuietKeys {
    bool disabled = false;
    HICON hIcon_mic;
    HICON hIcon_mic_fill;
    HICON hIcon_mic_mute;
    HICON hIcon_mic_mute_fill;
    NOTIFYICONDATA nid;
}

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_QUIETKEYSAPP, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    gQuietKeys::hIcon_mic = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC));
    gQuietKeys::hIcon_mic_fill = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_FILL));
    gQuietKeys::hIcon_mic_mute = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_MUTE));
    gQuietKeys::hIcon_mic_mute_fill = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_MUTE_FILL));

    // Perform application initialization:
    if (!InitInstance (hInstance, SW_HIDE))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_QUIETKEYSAPP));

    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}



//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_FILL));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_QUIETKEYSAPP);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_MIC_FILL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);
   
   // setup the notification icon
   gQuietKeys::nid = {};
   gQuietKeys::nid.cbSize = sizeof(gQuietKeys::nid);
   gQuietKeys::nid.hWnd = hWnd;
   gQuietKeys::nid.uID = 1;
   gQuietKeys::nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE;
   gQuietKeys::nid.uCallbackMessage = APPWM_ICONNOTIFY;
   gQuietKeys::nid.hIcon = gQuietKeys::hIcon_mic_fill;
   StringCchPrintf(gQuietKeys::nid.szTip,
       ARRAYSIZE(gQuietKeys::nid.szTip),
       L"%ws (Enabled)",
       szTitle);

   Shell_NotifyIcon(NIM_ADD, &gQuietKeys::nid);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case APPWM_ICONNOTIFY:
    {
        switch (lParam)
        {
        case WM_LBUTTONUP:
            // left-click to toggle whether enabled or disabled
            gQuietKeys::disabled = !gQuietKeys::disabled;
            if (gQuietKeys::disabled == TRUE) {
                gQuietKeys::nid.hIcon = gQuietKeys::hIcon_mic;
                StringCchPrintf(gQuietKeys::nid.szTip,
                    ARRAYSIZE(gQuietKeys::nid.szTip),
                    L"%ws (Disabled)",
                    szTitle);
            }
            else {
                gQuietKeys::nid.hIcon = gQuietKeys::hIcon_mic_fill;
                StringCchPrintf(gQuietKeys::nid.szTip,
                    ARRAYSIZE(gQuietKeys::nid.szTip),
                    L"%ws (Enabled)",
                    szTitle);
            }
            Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);
            break;
        case WM_RBUTTONUP:
            // right-click for the context menu
            SetForegroundWindow(hWnd);
            HMENU menu = LoadMenu(hInst, MAKEINTRESOURCEW(IDC_QUIETKEYSAPP));
            UINT flags = TPM_LEFTALIGN | TPM_TOPALIGN;
            POINT cursorpos;
            GetCursorPos(&cursorpos);
            BOOL menuVisible = TrackPopupMenuEx(
                GetSubMenu(menu, 0), flags, cursorpos.x, cursorpos.y, hWnd, NULL);
            break;
        }
    }
    break;
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

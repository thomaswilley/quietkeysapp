/*
* QuietKeys (QuietkeysApp.cpp et al)
* A tiny native win32/c++ app that mutes your default PC mic while you type (with a 2s delay when you're done)
* Copyright (c) Thomas Willey 2021
* Icons credit: https://getbootstrap.com
*/

#include "QuietkeysApp.h"
#include <shellapi.h>
#include <stdio.h>
#include <strsafe.h>
#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>
#include <commctrl.h> // NM_CLICK

// for using syslink ctl in about dialog. :/
// https://github.com/microsoftarchive/msdn-code-gallery-microsoft/blob/master/OneCodeTeam/Windows%20common%20controls%20demo%20(CppWindowsCommonControls)/%5BC%2B%2B%5D-Windows%20common%20controls%20demo%20(CppWindowsCommonControls)/C%2B%2B/CppWindowsCommonControls/CppWindowsCommonControls.cpp
#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define MAX_LOADSTRING      100
#define APPWM_ICONNOTIFY    (WM_APP + 1)
#define APPWM_MICMUTED      (WM_APP + 2)
#define APPWM_MICUNMUTED    (WM_APP + 3)

// for gQuietKeys::GetIcon
#define ICON_ENABLED        1
#define ICON_DISABLED       2
#define ICON_TYPING         3
#define ICON_DISABLED_ERROR 4

// forward declarations
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK LowLevelKeyboardProc(_In_ int nCode, _In_ WPARAM wParam, _In_ LPARAM lParam);
void TypingIsHappening();
void GetDefaultMic(wchar_t** defaultMicFriendlyName, IAudioEndpointVolume** defaultMicVolume);
BOOL IsDarkThemeActive();

// globals
HINSTANCE hInst;
WCHAR szTitle[MAX_LOADSTRING];
WCHAR szWindowClass[MAX_LOADSTRING];

namespace gQuietKeys {
    bool disabled = false;
    HICON hIcon_mic;
    HICON hIcon_mic_fill;
    HICON hIcon_mic_mute;
    HICON hIcon_mic_mute_fill;
    HICON hIcon_mic_white;
    HICON hIcon_mic_fill_white;
    HICON hIcon_mic_mute_white;
    HICON hIcon_mic_mute_fill_white;
    NOTIFYICONDATA nid;
    HWND hDlg_About;
    HHOOK hookKeyboard;
    wchar_t* defaultMicFriendlyName = NULL;
    wchar_t strFriendlyStatus[75];
    IAudioEndpointVolume* defaultMicVolume = NULL;
    bool mic_is_currently_muted = false; // this is used as a flag by the app, tbd replace later w/GetCurrentlyMuted()
    UINT_PTR QK_TIMER = 0x11; // gets updated as timer gets set.

    void SetFriendlyStatus(int icon_code)
    {
        switch (icon_code)
        {
        case ICON_DISABLED:
            StringCchPrintf(gQuietKeys::strFriendlyStatus, 75, L"%ws", L"QuietKeys is disabled. To enable, left-click its icon!");
            break;
        case ICON_ENABLED:
            StringCchPrintf(gQuietKeys::strFriendlyStatus, 75, L"Active on default mic: %ws", gQuietKeys::defaultMicFriendlyName);
            break;
        case ICON_TYPING:
            StringCchPrintf(gQuietKeys::strFriendlyStatus, 75, L"Active [now typing/muted] on default mic: %ws", gQuietKeys::defaultMicFriendlyName);
            break;
        case ICON_DISABLED_ERROR:
            StringCchPrintf(gQuietKeys::strFriendlyStatus, 75, L"%ws", L"QuietKeys is disabled: An error has occurred.");
            break;
        default:
            break;
        }
        if (gQuietKeys::hDlg_About != NULL)
        {
            PostMessage(gQuietKeys::hDlg_About, APPWM_ICONNOTIFY, NULL, NULL);
        }
    }

    HICON GetQKIcon(int icon_code) {
        BOOL bDark = IsDarkThemeActive();
        switch (icon_code) {
        case ICON_ENABLED:
            return bDark ? gQuietKeys::hIcon_mic_fill_white : gQuietKeys::hIcon_mic_fill;
            break;
        case ICON_TYPING:
            return bDark ? gQuietKeys::hIcon_mic_mute_fill_white : gQuietKeys::hIcon_mic_mute_fill;
            break;
        case ICON_DISABLED:
            return bDark ? gQuietKeys::hIcon_mic_white : gQuietKeys::hIcon_mic;
            break;
        default:
            return GetQKIcon(ICON_DISABLED);
        }
    }

    BOOL GetCurrentlyMuted()
    {
        BOOL muted;
        gQuietKeys::defaultMicVolume->GetMute(&muted);
        return muted;
    }

    void Unmute() {
        if (gQuietKeys::GetCurrentlyMuted() == TRUE) {
            gQuietKeys::defaultMicVolume->SetMute(FALSE, NULL);
            PostMessage(gQuietKeys::nid.hWnd, APPWM_MICUNMUTED, NULL, NULL);
            gQuietKeys::SetFriendlyStatus(ICON_ENABLED);
        }
    }

    void Mute() {
        if (gQuietKeys::GetCurrentlyMuted() == FALSE) {
            gQuietKeys::defaultMicVolume->SetMute(TRUE, NULL);
            PostMessage(gQuietKeys::nid.hWnd, APPWM_MICMUTED, NULL, NULL);
            gQuietKeys::SetFriendlyStatus(ICON_TYPING);
        }
    }

    float GetCurrentVolume()
    {
        float volume;
        gQuietKeys::defaultMicVolume->GetMasterVolumeLevelScalar(&volume);
        return volume;
    }

    void Timerproc(HWND Arg1, UINT Arg2, UINT_PTR Arg3, DWORD Arg4)
    {
        gQuietKeys::Unmute();
        KillTimer(NULL, gQuietKeys::QK_TIMER);
        gQuietKeys::mic_is_currently_muted = false;
    }
}

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
    gQuietKeys::hIcon_mic_mute = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_MUTE));
    gQuietKeys::hIcon_mic_mute_fill = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_MUTE_FILL));
    gQuietKeys::hIcon_mic_fill = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_FILL));
    gQuietKeys::hIcon_mic_white = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_WHITE));
    gQuietKeys::hIcon_mic_mute_white = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_MUTE_WHITE));
    gQuietKeys::hIcon_mic_mute_fill_white = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_MUTE_FILL_WHITE));
    gQuietKeys::hIcon_mic_fill_white = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_MIC_FILL_WHITE));
    
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
    wcex.hIcon          = gQuietKeys::GetQKIcon(ICON_ENABLED);
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_QUIETKEYSAPP);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = gQuietKeys::GetQKIcon(ICON_ENABLED);

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
   gQuietKeys::nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE | NIF_INFO;
   gQuietKeys::nid.uCallbackMessage = APPWM_ICONNOTIFY;
   gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_ENABLED);
   StringCchPrintf(gQuietKeys::nid.szTip,
       ARRAYSIZE(gQuietKeys::nid.szTip),
       L"%ws (Enabled)",
       szTitle);
   StringCchPrintf(gQuietKeys::nid.szInfoTitle,
       ARRAYSIZE(gQuietKeys::nid.szInfoTitle),
       L"%ws is running!",
       szTitle);
   StringCchPrintf(gQuietKeys::nid.szInfo,
       ARRAYSIZE(gQuietKeys::nid.szInfo),
       L"%ws",
       L"QuietKeys lives down in your system notification tray just like your battery % and wifi. Right-click it's icon to manage, left-click to enable/disable.");
   gQuietKeys::nid.uTimeout = 5000;
   Shell_NotifyIcon(NIM_ADD, &gQuietKeys::nid);

   gQuietKeys::nid.uFlags = NIF_ICON | NIF_TIP | NIF_MESSAGE; // invalidate NIF_INFO (so we don't show the balloon tooltip anymore.
   gQuietKeys::nid.hBalloonIcon = gQuietKeys::nid.hIcon;
   Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);

   // insert core keyboard hook
   gQuietKeys::hookKeyboard = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, NULL, NULL);
   // init com
   HRESULT hr = CoInitialize(NULL);
   if (hr != S_OK) {
       gQuietKeys::SetFriendlyStatus(ICON_DISABLED_ERROR);
       return FALSE;
   }
   // grab the default mic
   GetDefaultMic(&gQuietKeys::defaultMicFriendlyName, &gQuietKeys::defaultMicVolume);

   gQuietKeys::SetFriendlyStatus(ICON_ENABLED);

   return TRUE;
}

// a hacky way to see if the taskbar is dark or light themed so we can use the light / dark icon set
BOOL IsDarkThemeActive()
{
    DWORD   type;
    DWORD   value;
    DWORD   count = 4;
    LSTATUS st = RegGetValue(
        HKEY_CURRENT_USER,
        TEXT("Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize"),
        TEXT("SystemUsesLightTheme"),
        RRF_RT_REG_DWORD,
        &type,
        &value,
        &count);
    if (st == ERROR_SUCCESS && type == REG_DWORD)
        return (BOOL)value == 0;
    return FALSE;
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
    case APPWM_MICMUTED:
        gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_TYPING);
        Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);
        gQuietKeys::SetFriendlyStatus(ICON_TYPING);
        break;
    case APPWM_MICUNMUTED:
        gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_ENABLED);
        Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);
        gQuietKeys::SetFriendlyStatus(ICON_ENABLED);
        break;
    case APPWM_ICONNOTIFY:
    {
        switch (lParam)
        {
        case WM_LBUTTONUP:
            // left-click to toggle whether enabled or disabled
            gQuietKeys::disabled = !gQuietKeys::disabled;
            if (gQuietKeys::disabled == TRUE) {
                gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_DISABLED);
                StringCchPrintf(gQuietKeys::nid.szTip,
                    ARRAYSIZE(gQuietKeys::nid.szTip),
                    L"%ws (Disabled)",
                    szTitle);
                gQuietKeys::SetFriendlyStatus(ICON_DISABLED);
            }
            else {
                gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_ENABLED);
                StringCchPrintf(gQuietKeys::nid.szTip,
                    ARRAYSIZE(gQuietKeys::nid.szTip),
                    L"%ws (Enabled)",
                    szTitle);
                gQuietKeys::SetFriendlyStatus(ICON_ENABLED);
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
    case WM_SETTINGCHANGE:
    {
        if (!lstrcmp(LPCTSTR(lParam), L"ImmersiveColorSet"))
        {
            // poke the hwnd if dark/light them changes while running.
            if (gQuietKeys::disabled) {
                gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_DISABLED);
            }
            else {
                gQuietKeys::nid.hIcon = gQuietKeys::GetQKIcon(ICON_ENABLED);
            }
            Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);
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
                DialogBox(hInst, MAKEINTRESOURCEW(IDD_ABOUTBOX), hWnd, About);
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
        KillTimer(NULL, gQuietKeys::QK_TIMER);
        UnhookWindowsHookEx(gQuietKeys::hookKeyboard);
        gQuietKeys::defaultMicVolume->Release();
        gQuietKeys::defaultMicVolume = NULL;
        gQuietKeys::defaultMicFriendlyName = NULL;
        CoUninitialize();
        Shell_NotifyIcon(NIM_DELETE, &gQuietKeys::nid);
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
        gQuietKeys::hDlg_About = hDlg;
        SetDlgItemText(hDlg, IDC_ABOUTDLG_STATIC_STATUS, gQuietKeys::strFriendlyStatus);
        return (INT_PTR)TRUE;
        break;
    case APPWM_ICONNOTIFY:
        SetDlgItemText(hDlg, IDC_ABOUTDLG_STATIC_STATUS, gQuietKeys::strFriendlyStatus);
        break;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            gQuietKeys::hDlg_About = NULL;
            return (INT_PTR)TRUE;
        }
        SetDlgItemText(hDlg, IDC_ABOUTDLG_STATIC_STATUS, gQuietKeys::strFriendlyStatus);
        break;
    case WM_NOTIFY:
        switch (((LPNMHDR)lParam)->code)
        {
        case NM_CLICK:
            if (wParam == IDC_SYSLINK1)
            {
                PNMLINK pNMLink = (PNMLINK)lParam;
                LITEM   item = pNMLink->item;
                ShellExecute(NULL, L"open", item.szUrl, NULL, NULL, SW_SHOW);
            }
            break;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// win32 grab the audio input
void GetDefaultMic(wchar_t** defaultMicFriendlyName, IAudioEndpointVolume** defaultMicVolume)
{
    IMMDeviceEnumerator* deviceEnumerator = NULL;
    IMMDevice* defaultDevice = NULL;
    IPropertyStore* defaultDeviceProperties;
    PROPVARIANT deviceFriendlyName;

    HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator),
        NULL,
        CLSCTX_INPROC_SERVER,
        __uuidof(IMMDeviceEnumerator),
        (LPVOID*)&deviceEnumerator);

    hr = deviceEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &defaultDevice);
    hr = defaultDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_INPROC_SERVER, NULL, (LPVOID*)&*defaultMicVolume);
    hr = defaultDevice->OpenPropertyStore(STGM_READ, &defaultDeviceProperties);

    PropVariantInit(&deviceFriendlyName);
    hr = defaultDeviceProperties->GetValue(PKEY_DeviceInterface_FriendlyName, &deviceFriendlyName);

    int length = lstrlen((wchar_t*)deviceFriendlyName.pszVal);
    *defaultMicFriendlyName = new wchar_t[length];
    wsprintf(*defaultMicFriendlyName, (wchar_t*)deviceFriendlyName.pszVal);

    PropVariantClear(&deviceFriendlyName);
    deviceEnumerator->Release();
    deviceEnumerator = NULL;
    defaultDevice->Release();
    defaultDeviceProperties->Release();
    defaultDevice = NULL;
}

// the meat of the keyboard callback
void TypingIsHappening()
{
    if (gQuietKeys::defaultMicVolume == nullptr) {
        //OutputDebugString(L"typing is happening, but the mic isn't registered. try restarting the app.");
        return;
    }
    if (gQuietKeys::mic_is_currently_muted == false) {
        // mute the mic & set the timer.
        gQuietKeys::mic_is_currently_muted = true;
        gQuietKeys::Mute();
        gQuietKeys::QK_TIMER = SetTimer(NULL, gQuietKeys::QK_TIMER, 2000, (TIMERPROC)gQuietKeys::Timerproc);
    }
    // extend the timer while typing continues
    gQuietKeys::QK_TIMER = SetTimer(NULL, gQuietKeys::QK_TIMER, 2000, (TIMERPROC)gQuietKeys::Timerproc);
    // (typing is happening and mic is muted)
}

// the keyboard callback
LRESULT CALLBACK LowLevelKeyboardProc(_In_ int nCode, _In_ WPARAM wParam, _In_ LPARAM lParam)
{
    // https://msdn.microsoft.com/en-us/library/windows/desktop/ms644985(v=vs.85).aspx
    KBDLLHOOKSTRUCT* p = (KBDLLHOOKSTRUCT*)lParam;

    if (wParam == WM_KEYUP && !gQuietKeys::disabled) { // don't look at the actual key. just start muting (or stay muted) until done typing.
        TypingIsHappening();
    }
    return CallNextHookEx(NULL, nCode, wParam, lParam);
}
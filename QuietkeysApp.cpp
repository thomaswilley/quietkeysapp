// QuietkeysApp.cpp : Defines the entry point for the application.
//

#include "framework.h"
#include "QuietkeysApp.h"
#include <shellapi.h>
#include <stdio.h>
#include <strsafe.h>
#include <windows.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <Functiondiscoverykeys_devpkey.h>

#define MAX_LOADSTRING      100
#define APPWM_ICONNOTIFY    (WM_APP + 1)
#define APPWM_MICMUTED      (WM_APP + 2)
#define APPWM_MICUNMUTED    (WM_APP + 3)

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
    HHOOK hookKeyboard;
    wchar_t* defaultMicFriendlyName = NULL;
    IAudioEndpointVolume* defaultMicVolume = NULL;
    bool mic_is_currently_muted = false; // this is used as a flag by the app, tbd replace later w/GetCurrentlyMuted()
    UINT_PTR QK_TIMER = 0x11; // gets updated as timer gets set.
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
        }
    }
    void Mute() {
        if (gQuietKeys::GetCurrentlyMuted() == FALSE) {
            gQuietKeys::defaultMicVolume->SetMute(TRUE, NULL);
            PostMessage(gQuietKeys::nid.hWnd, APPWM_MICMUTED, NULL, NULL);
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
        OutputDebugString(L"typing is happening, but the mic isn't registered. try restarting the app.");
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

   // insert core keyboard hook
   gQuietKeys::hookKeyboard = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeyboardProc, NULL, NULL);
   // init com
   HRESULT hr = CoInitialize(NULL);
   if (hr != S_OK) {
       return FALSE;
   }
   // grab the default mic
   GetDefaultMic(&gQuietKeys::defaultMicFriendlyName, &gQuietKeys::defaultMicVolume);

   /*
   * TODO: update about dialog to display this
   * printf("Managing default mic: %ws (volume: %f, %ws)\n",
        defaultMicFriendlyName,
        quietkeys_global::GetCurrentVolume(),
        quietkeys_global::GetCurrentlyMuted() == TRUE ? L"currently muted" : L"currently unmuted");
   */

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
    case APPWM_MICMUTED:
        gQuietKeys::nid.hIcon = gQuietKeys::hIcon_mic_mute_fill;
        Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);
        break;
    case APPWM_MICUNMUTED:
        OutputDebugString(L"Unmute");
        gQuietKeys::nid.hIcon = gQuietKeys::hIcon_mic_fill;
        Shell_NotifyIcon(NIM_MODIFY, &gQuietKeys::nid);
        break;
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
        KillTimer(NULL, gQuietKeys::QK_TIMER);
        UnhookWindowsHookEx(gQuietKeys::hookKeyboard);
        gQuietKeys::defaultMicVolume->Release();
        gQuietKeys::defaultMicVolume = NULL;
        gQuietKeys::defaultMicFriendlyName = NULL;
        CoUninitialize();
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
        break;
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

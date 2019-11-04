#ifndef UTILS_H
#define UTILS_H

BOOL GetWorkAreaRect(HWND hwnd, __out RECT* prc) throw()
{
    HMONITOR hMonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
    if (hMonitor != NULL) {
        MONITORINFO mi = { sizeof(MONITORINFO) };
        BOOL bRet = GetMonitorInfo(hMonitor, &mi);
        if (bRet) {
            *prc = mi.rcWork;
            return bRet;
        }
    }
    return FALSE;
}

void CenterWindowPos(HWND hwnd, HWND hwndAfter, int width, int height, UINT flags = SWP_NOACTIVATE) throw()
{
    RECT parentRect;
    if (GetWorkAreaRect(hwnd, &parentRect)) {
        const LONG x = parentRect.left + ((parentRect.right - parentRect.left) - width) / 2;
        const LONG y = parentRect.top + ((parentRect.bottom - parentRect.top) - height) / 2;
        ::SetWindowPos(hwnd, hwndAfter, x, y, width, height, flags);
    }
}

void MyAdjustWindowRect(HWND hwnd, __inout LPRECT lprc) throw()
{
    ::AdjustWindowRect(lprc, WS_OVERLAPPEDWINDOW, TRUE);
    RECT rect = *lprc;
    rect.bottom = 0x7fff;
    FORWARD_WM_NCCALCSIZE(hwnd, FALSE, &rect, ::SendMessage);
    lprc->bottom += rect.top;
}

BOOL MyDragDetect(__in HWND hwnd, __in POINT pt) throw()
{
    const int cxDrag = ::GetSystemMetrics(SM_CXDRAG);
    const int cyDrag = ::GetSystemMetrics(SM_CYDRAG);
    const RECT dragRect = { pt.x - cxDrag, pt.y - cyDrag, pt.x + cxDrag, pt.y + cyDrag };

    // Start capture
    ::SetCapture(hwnd);

    do {
        // Sleep the current thread until any mouse/key input is in the queue.
        const DWORD status = ::MsgWaitForMultipleObjectsEx(0, NULL, INFINITE, QS_MOUSE | QS_KEY, MWMO_INPUTAVAILABLE);
        if (status != WAIT_OBJECT_0) {
            // Unexpected error.
            const DWORD dwLastError = ::GetLastError();
            ::ReleaseCapture();
            ::RestoreLastError(dwLastError);
            return FALSE;
        }

        MSG msg;
        while (::PeekMessage(&msg, NULL, 0, 0, PM_REMOVE)) {
            if (msg.message == WM_QUIT) {
                // Bail out due to WM_QUIT.
                ::ReleaseCapture();
                ::PostQuitMessage(static_cast<int>(msg.wParam));
                return FALSE;
            }

            // Give an application the opportunity to process this message first.
            if (::CallMsgFilter(&msg, MSGF_COMMCTRL_BEGINDRAG)) {
                continue;
            }

            switch (msg.message) {
            case WM_LBUTTONUP:
            case WM_RBUTTONUP:
            case WM_MBUTTONUP:
            case WM_XBUTTONUP:
            case WM_LBUTTONDOWN:
            case WM_RBUTTONDOWN:
            case WM_MBUTTONDOWN:
            case WM_XBUTTONDOWN:
                // A mouse button is pressed while the mouse cursor is staying inside the drag rectangle.
                ::ReleaseCapture();
                return FALSE;

            case WM_MOUSEMOVE:
                if (::PtInRect(&dragRect, msg.pt) == FALSE) {
                    // The mouse cursor is moved out of the drag rectangle.
                    ::ReleaseCapture();
                    return TRUE;
                }
                break;

            case WM_KEYDOWN:
                if (msg.wParam == VK_ESCAPE) {
                    // the ESC key is pressed.
                    ::ReleaseCapture();
                    return FALSE;
                }
                break;

            default:
                ::TranslateMessage(&msg);
                ::DispatchMessage(&msg);
                break;
            }
            break;
        }

        // Loop until the mouse capture is released by WM_CANCELMODE.
    } while (::GetCapture() == hwnd);

    return FALSE;
}

void SetWindowBlur(HWND hwnd)
{
    HINSTANCE hmod = ::LoadLibrary(TEXT("user32.dll"));
    if (hmod != NULL) {
        struct ACCENTPOLICY
        {
            int AccentState;
            int Flags;
            int Color;
            int AnimationId;
        };

        struct WINCOMPATTRDATA
        {
            int Attribute;
            PVOID pData;
            SIZE_T cbData;
        };

        // Might be broken in the future due to this undocumented API
        // https://www.google.com/search?q=SetWindowCompositionAttribute
        BOOL (WINAPI* SetWindowCompositionAttribute)(HWND, WINCOMPATTRDATA*);
        reinterpret_cast<FARPROC&>(SetWindowCompositionAttribute) = ::GetProcAddress(hmod, "SetWindowCompositionAttribute");
        if (SetWindowCompositionAttribute != NULL) {
            ACCENTPOLICY policy = { 3, 0, 0, 0 };
            WINCOMPATTRDATA data = { 19, &policy, sizeof(ACCENTPOLICY) };
            SetWindowCompositionAttribute(hwnd, &data);
        }
        ::FreeLibrary(hmod);
    }
}

#endif

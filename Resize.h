#ifndef RESIZE_H
#define RESIZE_H

template <class T>
class ResizableWindow
{
private:
    static const int SIZE_FRAME = 8;
    static const int SIZE_INNERFRAME = 16;

public:
    ResizableWindow() throw()
    {
    }

    ~ResizableWindow() throw()
    {
    }

    bool ProcessWindowMessages(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam, __out LRESULT* result) throw()
    {
        *result = 0;
        switch (message) {
        case WM_SETCURSOR:
            return this->OnSetCursor(hwnd, wParam, lParam, result);
        case WM_NCHITTEST:
            return this->OnNcHitTest(hwnd, wParam, lParam, result);
        case WM_NCLBUTTONDOWN:
            return this->OnNcLButtonDown(hwnd, wParam, lParam, result);
        }
        return false;
    }

private:
    bool OnSetCursor(HWND, WPARAM, LPARAM lParam, __out_opt LRESULT* result) throw()
    {
        UINT htCode = LOWORD(lParam);
#define IDCtoCUR(idc) (LOBYTE(idc))
#define CURtoIDC(cur) MAKEINTRESOURCE((cur) | 0x7F00)
        static const BYTE htsizCurs[HTSIZELAST - HTSIZEFIRST + 1] =
        {
            IDCtoCUR(IDC_SIZEWE),   // HTLEFT
            IDCtoCUR(IDC_SIZEWE),   // HTRIGHT
            IDCtoCUR(IDC_SIZENS),   // HTTOP
            IDCtoCUR(IDC_SIZENWSE), // HTTOPLEFT
            IDCtoCUR(IDC_SIZENESW), // HTTOPRIGHT
            IDCtoCUR(IDC_SIZENS),   // HTBOTTOM
            IDCtoCUR(IDC_SIZENESW), // HTBOTTOMLEFT
            IDCtoCUR(IDC_SIZENWSE), // HTBOTTOMRIGHT
        };

        LPCTSTR cursor = IDC_ARROW;
        if (static_cast<UINT>(htCode - HTSIZEFIRST) <= (HTSIZELAST - HTSIZEFIRST)) {
            cursor = CURtoIDC(htsizCurs[htCode - HTSIZEFIRST]);
        }
        ::SetCursor(::LoadCursor(NULL, cursor));
#undef IDCtoCUR
#undef CURtoIDC
        *result = TRUE;
        return true;
    }

    bool OnNcHitTest(HWND hwnd, WPARAM, LPARAM lParam, __out_opt LRESULT* result) throw()
    {
        RECT outerRect = {};
        ::GetWindowRect(hwnd, &outerRect);
        RECT innerRect = outerRect;
        ::InflateRect(&outerRect, -SIZE_FRAME, -SIZE_FRAME);
        ::InflateRect(&innerRect, -SIZE_INNERFRAME, -SIZE_INNERFRAME);

        static const signed char hitTests[3][5] =
        {
            { HTTOPLEFT,    HTTOPLEFT,    HTTOP,    HTTOPRIGHT,    HTTOPRIGHT    },
            { HTLEFT,       HTCLIENT,     HTCLIENT, HTCLIENT,      HTRIGHT       },
            { HTBOTTOMLEFT, HTBOTTOMLEFT, HTBOTTOM, HTBOTTOMRIGHT, HTBOTTOMRIGHT },
        };

        int x = GET_X_LPARAM(lParam);
        int y = GET_Y_LPARAM(lParam);
        UINT col = (x >= outerRect.left) + (x >= innerRect.left) + (x >= innerRect.right) + (x >= outerRect.right);
        UINT row = (y >= outerRect.top)  + (y >= outerRect.bottom);
        *result = hitTests[row][col];
        return true;
    }

    bool OnNcLButtonDown(HWND hwnd, WPARAM wParam, LPARAM lParam, __out_opt LRESULT*) throw()
    {
        UINT htCode = static_cast<UINT>(wParam);
        UINT idCmd;
        if (static_cast<UINT>(htCode - HTSIZEFIRST) <= (HTSIZELAST - HTSIZEFIRST)) {
            idCmd = SC_SIZE + (htCode - HTSIZEFIRST + 1);
            ::SendMessage(hwnd, WM_SYSCOMMAND, idCmd, lParam);
            return true;
        }
        return false;
    }
};

#endif


class MagneticWindow
{
private:
    POINT hotspot;
    bool dragging;
    static const UINT SNAP_PIXELS = 20;

    void AdjustPoint(HWND hwnd, __inout POINT* pt)
    {
        if (::GetAsyncKeyState(VK_SHIFT) >= 0) {
            RECT workRect = {};
            GetWorkAreaRect(hwnd, &workRect);
            RECT rect = {};
            ::GetWindowRect(hwnd, &rect);
            pt->x = DoSnap(pt->x, workRect.left);
            pt->y = DoSnap(pt->y, workRect.top);
            pt->x = DoSnap(pt->x, workRect.right - (rect.right - rect.left));
            pt->y = DoSnap(pt->y, workRect.bottom - (rect.bottom - rect.top));
        }
    }

    static LONG DoSnap(LONG a, LONG b)
    {
        return (static_cast<ULONG>(::labs(b - a)) < SNAP_PIXELS) ? b : a;
    }

public:
    MagneticWindow() throw()
        : hotspot()
        , dragging()
    {
    }

    ~MagneticWindow() throw()
    {
    }

    void BeginDrag(HWND hwnd, int x, int y) throw()
    {
        this->hotspot.x = x;
        this->hotspot.y = y;
        ::SetCapture(hwnd);
        this->dragging = true;
    }

    void EndDrag() throw()
    {
        this->dragging = false;
        ::ReleaseCapture();
    }

    void Drag(HWND hwnd, int x, int y) throw()
    {
        if (this->dragging && ::GetCapture() == hwnd) {
            POINT pt = {
                x - this->hotspot.x,
                y - this->hotspot.y
            };
            ::ClientToScreen(hwnd, &pt);
            this->AdjustPoint(hwnd, &pt);
            ::SetWindowPos(hwnd, NULL, pt.x, pt.y, 0, 0, SWP_NOSIZE);
        }
    }
};

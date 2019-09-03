#define WINVER          0x0601
#define _WIN32_WINNT    0x0601
#define _WIN32_IE       0x0800
#define UNICODE
#define _UNICODE

#include <sdkddkver.h>
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <wincodec.h>
#include <uxtheme.h>
#include <stddef.h>
#include <intrin.h>
#include <math.h>

#include <strsafe.h>

WINCOMMCTRLAPI HRESULT WINAPI ImageList_CoCreateInstance(__in REFCLSID rclsid, __in_opt const IUnknown *punkOuter, __in REFIID riid, __deref_out void **ppv);

#undef INTERFACE
#define INTERFACE IImageList

DECLARE_INTERFACE_IID_(IImageList, IUnknown, "46EB5926-582E-4017-9FDF-E8998DAA0950")
{
    BEGIN_INTERFACE

    STDMETHOD(QueryInterface)(THIS_ REFIID riid, void **ppv) PURE;
    STDMETHOD_(ULONG,AddRef)(THIS) PURE;
    STDMETHOD_(ULONG,Release)(THIS) PURE;
    STDMETHOD(Add)(THIS_ __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask, __out int *pi) PURE;
    STDMETHOD(ReplaceIcon)(THIS_ int i, __in HICON hicon, __out int *pi) PURE;
    STDMETHOD(SetOverlayImage)(THIS_ int iImage, int iOverlay) PURE;
    STDMETHOD(Replace)(THIS_ int i, __in HBITMAP hbmImage, __in_opt HBITMAP hbmMask) PURE;
    STDMETHOD(AddMasked)(THIS_ __in HBITMAP hbmImage, COLORREF crMask, __out int *pi) PURE;
    STDMETHOD(Draw)(THIS_ __in IMAGELISTDRAWPARAMS *pimldp) PURE;
    STDMETHOD(Remove)(THIS_ int i) PURE;
    STDMETHOD(GetIcon)(THIS_ int i, UINT flags, __out HICON *picon) PURE;
    STDMETHOD(GetImageInfo)(THIS_ int i, __out IMAGEINFO *pImageInfo) PURE;
    STDMETHOD(Copy)(THIS_ int iDst, __in IUnknown *punkSrc, int iSrc, UINT uFlags) PURE;
    STDMETHOD(Merge)(THIS_ int i1, __in IUnknown *punk2, int i2, int dx, int dy, REFIID riid, __deref_out void **ppv) PURE;
    STDMETHOD(Clone)(THIS_ REFIID riid, __deref_out void **ppv) PURE;
    STDMETHOD(GetImageRect)(THIS_ int i, __out RECT *prc) PURE;
    STDMETHOD(GetIconSize)(THIS_ __out int *cx, __out int *cy) PURE;
    STDMETHOD(SetIconSize)(THIS_ int cx, int cy) PURE;
    STDMETHOD(GetImageCount)(THIS_ __out int *pi) PURE;
    STDMETHOD(SetImageCount)(THIS_ UINT uNewCount) PURE;
    STDMETHOD(SetBkColor)(THIS_ COLORREF clrBk, __out COLORREF *pclr) PURE;
    STDMETHOD(GetBkColor)(THIS_ __out COLORREF *pclr) PURE;
    STDMETHOD(BeginDrag)(THIS_ int iTrack, int dxHotspot, int dyHotspot) PURE;
    STDMETHOD(EndDrag)(THIS_ void) PURE;
    STDMETHOD(DragEnter)(THIS_ __in_opt HWND hwndLock, int x, int y) PURE;
    STDMETHOD(DragLeave)(THIS_ __in_opt HWND hwndLock) PURE;
    STDMETHOD(DragMove)(THIS_ int x, int y) PURE;
    STDMETHOD(SetDragCursorImage)(THIS_ __in IUnknown *punk, int iDrag, int dxHotspot, int dyHotspot) PURE;
    STDMETHOD(DragShowNolock)(THIS_ BOOL fShow) PURE;
    STDMETHOD(GetDragImage)(THIS_ __out_opt POINT *ppt, __out_opt POINT *pptHotspot, REFIID riid, __deref_out void **ppv) PURE;
    STDMETHOD(GetItemFlags)(THIS_ int i, __out DWORD *dwFlags) PURE;
    STDMETHOD(GetOverlayImage)(THIS_ int iOverlay, __out int *piIndex) PURE;

    END_INTERFACE
};

//#define USETRACE

// https://blogs.msdn.microsoft.com/oldnewthing/20041025-00/?p=37483/
extern "C" IMAGE_DOS_HEADER     __ImageBase;
#define HINST_THISCOMPONENT     ((HINSTANCE)(&__ImageBase))

/* void Cls_OnSettingChange(HWND hwnd, UINT uiParam, LPCTSTR lpszSectionName) */
#define HANDLE_WM_SETTINGCHANGE(hwnd, wParam, lParam, fn) \
    ((fn)((hwnd), (UINT)(wParam), (LPCTSTR)(lParam)), 0L)
#define FORWARD_WM_SETTINGCHANGE(hwnd, uiParam, lpszSectionName, fn) \
    (void)(fn)((hwnd), WM_WININICHANGE, (WPARAM)(UINT)(uiParam), (LPARAM)(LPCTSTR)(lpszSectionName))

// HANDLE_WM_CONTEXTMENU is buggy
// https://blogs.msdn.microsoft.com/oldnewthing/20040921-00/?p=37813
/* void Cls_OnContextMenu(HWND hwnd, HWND hwndContext, int xPos, int yPos) */
#undef HANDLE_WM_CONTEXTMENU
#undef FORWARD_WM_CONTEXTMENU
#define HANDLE_WM_CONTEXTMENU(hwnd, wParam, lParam, fn) \
    ((fn)((hwnd), (HWND)(wParam), (int)GET_X_LPARAM(lParam), (int)GET_Y_LPARAM(lParam)), 0L)
#define FORWARD_WM_CONTEXTMENU(hwnd, hwndContext, xPos, yPos, fn) \
    (void)(fn)((hwnd), WM_CONTEXTMENU, (WPARAM)(HWND)(hwndContext), MAKELPARAM((int)(xPos), (int)(yPos)))

#define FAIL_HR(hr) if (FAILED((hr))) goto Fail

template <class T>
void IUnknown_SafeAssign(T*& dst, T* src) throw()
{
    if (src != NULL) {
        src->AddRef();
    }
    if (dst != NULL) {
        dst->Release();
    }
    dst = src;
}

template <class T>
void IUnknown_SafeRelease(T*& unknown) throw()
{
    if (unknown != NULL) {
        unknown->Release();
        unknown = NULL;
    }
}

void Trace(__in __format_string const char* format, ...)
{
    TCHAR buffer[512];
    va_list argPtr;
    va_start(argPtr, format);
#ifdef _UNICODE
    WCHAR wformat[512];
    MultiByteToWideChar(CP_ACP, 0, format, -1, wformat, _countof(wformat));
    StringCchVPrintfW(buffer, _countof(buffer), wformat, argPtr);
#else
    StringCchVPrintfA(buffer, _countof(buffer), format, argPtr);
#endif
    WriteConsole(GetStdHandle(STD_OUTPUT_HANDLE), buffer, lstrlen(buffer), NULL, NULL);
    va_end(argPtr);
}

#undef TRACE
#undef ASSERT

#ifdef USETRACE
#define TRACE(s, ...)       Trace(s, __VA_ARGS__)
#define ASSERT(e)           do if (!(e)) { TRACE("%s(%d): Assertion failed\n", __FILE__, __LINE__); if (::IsDebuggerPresent()) { ::DebugBreak(); } } while(0)
#else
#define TRACE(s, ...)
#define ASSERT(e)
#endif

inline void* operator new(size_t size) throw()
{
    return ::malloc(size);
}

inline void* operator new[](size_t size) throw()
{
    return ::malloc(size);
}

inline void operator delete(void* ptr) throw()
{
    ::free(ptr);
}

inline void operator delete[](void* ptr) throw()
{
    ::free(ptr);
}

template <typename T>
inline HRESULT SafeAlloc(size_t size, __deref_out_bcount(size) T*& ptr) throw()
{
    ptr = static_cast<T*>(::malloc(size));
    return ptr != NULL ? S_OK : E_OUTOFMEMORY;
}

template <typename T>
inline void SafeFree(__deallocate_opt(Mem) T*& ptr) throw()
{
    ::free(ptr);
    ptr = NULL;
}

template <typename T>
inline HRESULT SafeAllocAligned(size_t size, __deref_out_bcount(size) T*& ptr) throw()
{
    ptr = static_cast<T*>(::_aligned_malloc(size, 16));
    return ptr != NULL ? S_OK : E_OUTOFMEMORY;
}

template <typename T>
inline void SafeFreeAligned(__deallocate_opt(Mem) T*& ptr) throw()
{
    ::_aligned_free(ptr);
    ptr = NULL;
}

template <class T>
inline HRESULT SafeNew(__deref_out T*& ptr) throw()
{
    ptr = new T();
    return ptr != NULL ? S_OK : E_OUTOFMEMORY;
}

template <class T>
inline void SafeDelete(__deallocate_opt(Mem) T*& ptr) throw()
{
    delete ptr;
    ptr = NULL;
}

inline HRESULT SafeStrDupW(const wchar_t* src, wchar_t*& dst) throw()
{
    dst = ::_wcsdup(src);
    return dst != NULL ? S_OK : E_OUTOFMEMORY;
}


template <class T>
class ComPtr
{
private:
    mutable T* ptr;

public:
    ComPtr() throw()
        : ptr()
    {
    }
    ~ComPtr() throw()
    {
        this->Release();
    }
    ComPtr(const ComPtr<T>& src) throw()
        : ptr()
    {
        operator =(src);
    }
    T* Detach() throw()
    {
        T* ptr = this->ptr;
        this->ptr = NULL;
        return ptr;
    }
    void Release() throw()
    {
        IUnknown_SafeRelease(this->ptr);
    }
    ComPtr<T>& operator =(T* src) throw()
    {
        IUnknown_SafeAssign(this->ptr, src);
        return *this;
    }
    operator T*() const throw()
    {
        return this->ptr;
    }
    T* operator ->() const throw()
    {
        return this->ptr;
    }
    T** operator &() throw()
    {
        ASSERT(this->ptr == NULL);
        return &this->ptr;
    }
    void CopyTo(__deref_out T** outPtr) throw()
    {
        ASSERT(this->ptr != NULL);
        (*outPtr = this->ptr)->AddRef();
    }
    HRESULT CoCreateInstance(REFCLSID clsid, IUnknown* outer, DWORD context) throw()
    {
        ASSERT(this->ptr == NULL);
        return ::CoCreateInstance(clsid, outer, context, IID_PPV_ARGS(&this->ptr));
    }
    template <class Q>
    HRESULT QueryInterface(__deref_out Q** outPtr) throw()
    {
        return this->ptr->QueryInterface(IID_PPV_ARGS(outPtr));
    }
};


static HRESULT StreamRead(IStream* stream, void* data, DWORD size)
{
    if (size == 0) {
        return S_OK;
    }
    DWORD read = 0;
    HRESULT hr = stream->Read(data, size, &read);
    if (SUCCEEDED(hr) && size != read) {
        hr = E_FAIL;
    }
    return hr;
}

static HRESULT StreamWrite(IStream* stream, const void* data, DWORD size)
{
    if (size == 0) {
        return S_OK;
    }
    DWORD written = 0;
    HRESULT hr = stream->Write(data, size, &written);
    if (SUCCEEDED(hr) && size != written) {
        hr = E_FAIL;
    }
    return hr;
}


enum ControlID
{
    IDC_LISTVIEW = 1,
    IDC_STATUSBAR,
    IDC_EDIT,
    IDC_LAST,
};

HBITMAP Create32bppTopDownDIB(int cx, int cy, __deref_out_opt void** bits) throw()
{
    BITMAPINFOHEADER bmi = {
        sizeof(BITMAPINFOHEADER),
        cx, -cy, 1, 32
    };
    return ::CreateDIBSection(NULL, reinterpret_cast<BITMAPINFO*>(&bmi), DIB_RGB_COLORS, bits, NULL, 0);
}

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

void SetWindowPosition(HWND hwnd, int x, int y, int cx, int cy) throw()
{
    RECT rcWork;
    if (GetWorkAreaRect(hwnd, &rcWork)) {
        int dx = 0, dy = 0;
        if (x < rcWork.left) {
            dx = rcWork.left - x;
        }
        if (y < rcWork.top) {
            dy = rcWork.top - y;
        }
        if (x + cx + dx > rcWork.right) {
            dx = rcWork.right - x - cx;
            if (x + dx < rcWork.left) {
                dx += rcWork.left - x;
            }
        }
        if (y + cy + dy > rcWork.bottom) {
            dy = rcWork.bottom - y - cy;
            if (y + dx < rcWork.top) {
                dy += rcWork.top - y;
            }
        }
        ::MoveWindow(hwnd, x + dx, y + dy, cx, cy, TRUE);
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

HRESULT BrowseForFolder(__in_opt HWND hwndOwner, __deref_out LPWSTR* path) throw()
{
    *path = NULL;
    HRESULT hr;
    ComPtr<IFileOpenDialog> dlg;
    ComPtr<IShellItem> si;
    DWORD options;

    FAIL_HR(hr = dlg.CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_INPROC_SERVER));
    FAIL_HR(hr = dlg->GetOptions(&options));
    FAIL_HR(hr = dlg->SetOptions(options | FOS_FORCEFILESYSTEM | FOS_PICKFOLDERS));
    FAIL_HR(hr = dlg->Show(hwndOwner));
    FAIL_HR(hr = dlg->GetResult(&si));

    FAIL_HR(hr = si->GetDisplayName(SIGDN_FILESYSPATH, path));

Fail:
    return hr;
}


template <class T>
class BaseWindow
{
private:
    static LRESULT CALLBACK sWindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) throw()
    {
        LPCREATESTRUCT lpcs = reinterpret_cast<LPCREATESTRUCT>(lParam);
        T* self;
        if (message == WM_NCCREATE && lpcs != NULL) {
            self = static_cast<T*>(lpcs->lpCreateParams);
            ::SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        } else {
            self = reinterpret_cast<T*>(::GetWindowLongPtr(hwnd, GWLP_USERDATA));
        }
        if (self != NULL) {
            return self->WindowProc(hwnd, message, wParam, lParam);
        } else {
            return ::DefWindowProc(hwnd, message, wParam, lParam);
        }
    }

protected:
    static ATOM Register(LPCTSTR className, LPCTSTR menu = NULL) throw()
    {
        WNDCLASS wc = {
            0, sWindowProc, 0, 0,
            HINST_THISCOMPONENT,
            NULL, ::LoadCursor(NULL, IDC_ARROW),
            0, menu,
            className };
        return ::RegisterClass(&wc);
    }

    HWND Create(ATOM atom, LPCTSTR name, DWORD style = WS_OVERLAPPEDWINDOW, DWORD exStyle = 0) throw()
    {
        return ::CreateWindowEx(exStyle, MAKEINTATOM(atom), name,
            style,
            CW_USEDEFAULT, CW_USEDEFAULT,
            CW_USEDEFAULT, CW_USEDEFAULT,
            NULL, NULL, HINST_THISCOMPONENT, static_cast<T*>(this));
    }
};


template <class T>
class DECLSPEC_NOVTABLE DropTargetImpl : public IDropTarget, public IOleWindow
{
protected:
    ComPtr<IDataObject> dataObjectDragging;

    typedef DWORD DROPEFFECT;

private:
    HWND hwnd;

public:
    DropTargetImpl() throw()
        : hwnd()
    {
    }

    ~DropTargetImpl() throw()
    {
    }

    STDMETHOD(QueryInterface)(REFIID iid, void** ppv) throw()
    {
        if (IsEqualIID(iid, __uuidof(IUnknown)) || IsEqualIID(iid, __uuidof(IDropTarget))) {
            *ppv = static_cast<IDropTarget*>(this);
            this->AddRef();
            return S_OK;
        }
        if (IsEqualIID(iid, __uuidof(IOleWindow))) {
            *ppv = static_cast<IOleWindow*>(this);
            this->AddRef();
        }

        *ppv = NULL;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() throw()
    {
        return 2;
    }

    STDMETHOD_(ULONG, Release)() throw()
    {
        return 1;
    }

    STDMETHOD(DragEnter)(IDataObject* dataObject, DWORD grfKeyState, POINTL ptl, __inout DWORD* effect) throw()
    {
        TRACE("IDropTarget::DragEnter, Effect = %d\n", *effect);

        this->dataObjectDragging = dataObject;

        POINT pt = { ptl.x, ptl.y };
        T* self = static_cast<T*>(this);

        self->ScreenToClient(&pt);
        *effect = self->OnDragEnter(dataObject, grfKeyState, pt);

        return S_OK;
    }

    STDMETHOD(DragOver)(DWORD grfKeyState, POINTL ptl, __inout DWORD* effect) throw()
    {
        TRACE("IDropTarget::DragOver, Effect = %d\n", *effect);

        POINT pt = { ptl.x, ptl.y };
        T* self = static_cast<T*>(this);

        self->ScreenToClient(&pt);
        *effect = self->OnDragOver(grfKeyState, pt);

        return S_OK;
    }

    STDMETHOD(DragLeave)() throw()
    {
        TRACE("IDropTarget::DragLeave\n");

        T* self = static_cast<T*>(this);
        self->OnDragLeave();

        this->dataObjectDragging.Release();
        return S_OK;
    }

    STDMETHOD(Drop)(IDataObject* dataObject, DWORD grfKeyState, POINTL ptl, __inout DWORD* effect) throw()
    {
        TRACE("IDropTarget::Drop\n");

        POINT pt = { ptl.x, ptl.y };
        T* self = static_cast<T*>(this);

        self->ScreenToClient(&pt);
        *effect = self->OnDrop(dataObject, grfKeyState, pt);
        this->dataObjectDragging.Release();
        return S_OK;
    }

    STDMETHOD(GetWindow)(HWND* phwnd) throw()
    {
        if ((*phwnd = this->hwnd) != NULL) {
            return S_OK;
        }
        return E_FAIL;
    }

    STDMETHOD(ContextSensitiveHelp)(BOOL) throw()
    {
        return E_NOTIMPL;
    }

    HRESULT RegisterDragDrop(HWND hwnd) throw()
    {
        if (this->hwnd != NULL) {
            return S_FALSE;
        }
        IDropTarget* self = static_cast<IDropTarget*>(this);
        HRESULT hr = ::CoLockObjectExternal(self, TRUE, FALSE);
        if (SUCCEEDED(hr)) {
            hr = ::RegisterDragDrop(hwnd, self);
            if (SUCCEEDED(hr)) {
                this->hwnd = hwnd;
                return hr;
            }
            ::CoLockObjectExternal(self, FALSE, FALSE);
        }
        return hr;
    }

    HRESULT RevokeDragDrop() throw()
    {
        if (this->hwnd == NULL) {
            return S_FALSE;
        }
        IDropTarget* self = static_cast<IDropTarget*>(this);
        HRESULT hr = ::RevokeDragDrop(this->hwnd);
        if (SUCCEEDED(hr)) {
            ::CoLockObjectExternal(self, FALSE, TRUE);
            this->hwnd = NULL;
        }
        return hr;
    }

    HRESULT QueryGetDataObject(IDataObject* dataObject, CLIPFORMAT cfFormat) throw()
    {
        FORMATETC fe = { cfFormat, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        return dataObject->QueryGetData(&fe);
    }

    // overrides
    DROPEFFECT OnDragEnter(IDataObject* dataObject, DWORD grfKeyState, POINT pt) throw()
    {
        return DROPEFFECT_NONE;
    }

    DROPEFFECT OnDragOver(DWORD grfKeyState, POINT pt) throw()
    {
        return DROPEFFECT_NONE;
    }

    void OnDragLeave() throw()
    {
    }

    DROPEFFECT OnDrop(IDataObject* dataObject, DWORD grfKeyState, POINT pt) throw()
    {
        return DROPEFFECT_NONE;
    }

    void ScreenToClient(__inout LPPOINT ppt) throw()
    {
        ::ScreenToClient(this->hwnd, ppt);
    }
};


template <class T>
class DECLSPEC_NOVTABLE DropTargetExImpl : public DropTargetImpl<T>
{
protected:
    ComPtr<IDropTargetHelper> helper;
    bool showImage;

public:
    DropTargetExImpl() throw()
        : showImage(true)
    {
        ::CoCreateInstance(CLSID_DragDropHelper, 0, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&this->helper));
    }

    ~DropTargetExImpl() throw()
    {
    }

    STDMETHOD(DragEnter)(IDataObject* dataObject, DWORD grfKeyState, POINTL ptl, __inout DWORD* effect) throw()
    {
        TRACE("IDropTarget::DragEnter, Effect = %d\n", *effect);

        this->dataObjectDragging = dataObject;

        POINT pt = { ptl.x, ptl.y };
        T* self = static_cast<T*>(this);

        self->ScreenToClient(&pt);
        *effect = self->OnDragEnter(dataObject, grfKeyState, pt);

        if (this->showImage && this->helper != NULL) {
            HWND hwnd = NULL;
            this->GetWindow(&hwnd);
            pt.x = ptl.x;
            pt.y = ptl.y;
            this->helper->DragEnter(hwnd, dataObject, &pt, *effect);
        }
        return S_OK;
    }

    STDMETHOD(DragOver)(DWORD grfKeyState, POINTL ptl, __inout DWORD* effect) throw()
    {
        TRACE("IDropTarget::DragOver, Effect = %d\n", *effect);

        POINT pt = { ptl.x, ptl.y };
        T* self = static_cast<T*>(this);

        self->ScreenToClient(&pt);
        *effect = self->OnDragOver(grfKeyState, pt);

        if (this->showImage && this->helper != NULL) {
            pt.x = ptl.x;
            pt.y = ptl.y;
            this->helper->DragOver(&pt, *effect);
        }
        return S_OK;
    }

    STDMETHOD(DragLeave)() throw()
    {
        TRACE("IDropTarget::DragLeave\n");

        T* self = static_cast<T*>(this);
        self->OnDragLeave();
        this->dataObjectDragging.Release();

        if (this->showImage && this->helper != NULL) {
            this->helper->DragLeave();
        }
        return S_OK;
    }

    STDMETHOD(Drop)(IDataObject* dataObject, DWORD grfKeyState, POINTL ptl, __inout DWORD* effect) throw()
    {
        TRACE("IDropTarget::Drop\n");

        POINT pt = { ptl.x, ptl.y };
        T* self = static_cast<T*>(this);

        self->ScreenToClient(&pt);
        *effect = self->OnDrop(dataObject, grfKeyState, pt);

        if (this->showImage && this->helper != NULL) {
            pt.x = ptl.x;
            pt.y = ptl.y;
            this->helper->Drop(this->dataObjectDragging, &pt, *effect);
        }

        this->dataObjectDragging.Release();
        return S_OK;
    }

    void ShowDragImage(BOOL show) throw()
    {
        this->showImage = show != FALSE;
        if (this->helper != NULL) {
            this->helper->Show(show);
        }
    }
};


class MainWindow : public BaseWindow<MainWindow>, public DropTargetExImpl<MainWindow>
{
protected:
    friend BaseWindow<MainWindow>;
    friend DropTargetImpl<MainWindow>;
    friend DropTargetExImpl<MainWindow>;

    HBITMAP bitmap;
    int width;
    int height;
    ComPtr<IImageList> imageList;
    ComPtr<IDataObject> dataObject;

    static void ErrorMessage(HWND hwnd, LPCWSTR name) throw()
    {
        WCHAR message[MAX_PATH + 64];
        ::StringCchCopyW(message, _countof(message), L"Could not copy:\n");
        ::StringCchCatW(message, _countof(message), name);
        ::MessageBoxW(hwnd, message, L"Sorry", MB_OK);
    }

    LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) throw()
    {
        switch (message) {
        HANDLE_MSG(hwnd, WM_CREATE, OnCreate);
        HANDLE_MSG(hwnd, WM_DESTROY, OnDestroy);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_WINDOWPOSCHANGED, OnWindowPosChanged);
        HANDLE_MSG(hwnd, WM_SETTINGCHANGE, OnSettingChange);
        HANDLE_MSG(hwnd, WM_ACTIVATE, OnActivate);
        HANDLE_MSG(hwnd, WM_NOTIFY, OnNotify);
        HANDLE_MSG(hwnd, WM_INITMENU, OnInitMenu);
        HANDLE_MSG(hwnd, WM_INITMENUPOPUP, OnInitMenuPopup);
        HANDLE_MSG(hwnd, WM_CONTEXTMENU, OnContextMenu);
        HANDLE_MSG(hwnd, WM_LBUTTONDOWN, OnLButtonDown);
        HANDLE_MSG(hwnd, WM_PAINT, OnPaint);
        }
        return ::DefWindowProc(hwnd, message, wParam, lParam);
    }

    bool OnCreate(HWND hwnd, LPCREATESTRUCT lpcs) throw()
    {
        if (lpcs == NULL) {
            return false;
        }
        this->RegisterDragDrop(hwnd);
        return true;
    }

    void OnDestroy(HWND) throw()
    {
        this->dataObject.Release();
        this->RevokeDragDrop();
        ::PostQuitMessage(0);
    }

    void OnCommand(HWND hwnd, int id, HWND, UINT code) throw()
    {
    }

    void OnWindowPosChanged(HWND hwnd, const WINDOWPOS* pwp) throw()
    {
        if (pwp == NULL) {
            return;
        }
        // if (pwp->flags & SWP_SHOWWINDOW) {
            // if (this->shown == false) {
                // this->shown = true;
            // }
        // }
        if ((pwp->flags & SWP_NOSIZE) == 0) {
            ::InvalidateRect(hwnd, NULL, FALSE);
        }
    }

    void OnSettingChange(HWND hwnd, UINT, LPCTSTR) throw()
    {
        ::InvalidateRect(hwnd, NULL, FALSE);
    }

    void OnActivate(HWND, UINT state, HWND, BOOL) throw()
    {
    }

    LRESULT OnNotify(HWND hwnd, int idFrom, NMHDR* pnmhdr) throw()
    {
        return 0;
    }

    void OnInitMenu(HWND, HMENU hMenu) throw()
    {
    }

    void OnInitMenuPopup(HWND, HMENU hMenu, UINT, BOOL) throw()
    {
    }

    void OnContextMenu(HWND hwnd, HWND hwndContext, int xPos, int yPos) throw()
    {
        FORWARD_WM_CONTEXTMENU(hwnd, hwndContext, xPos, yPos, ::DefWindowProc);
    }

    void OnLButtonDown(HWND hwnd, BOOL, int, int, UINT) throw()
    {
        if (this->dataObject == NULL) {
            return;
        }
        POINT pt = {};
        ::GetCursorPos(&pt);
        if (::DragDetect(hwnd, pt)) {
            DWORD effect = DROPEFFECT_COPY;
            ::SHDoDragDrop(hwnd, this->dataObject, NULL, effect, &effect);
        }
    }

    void OnPaint(HWND hwnd) throw()
    {
        PAINTSTRUCT ps = {};
        HDC hdc = ::BeginPaint(hwnd, &ps);
        if (hdc != NULL) {
            RECT rect = {};
            ::GetClientRect(hwnd, &rect);
            ::PatBlt(hdc, 0, 0, rect.right, rect.bottom, WHITENESS);
            if (this->bitmap != NULL) {
                int width, height;
                double scale = this->ScaleSize(rect.right, rect.bottom, &width, &height);
                int x = (rect.right - width) / 2;
                int y = (rect.bottom - height) / 2;
                HDC hdcMem = ::CreateCompatibleDC(hdc);
                HBITMAP bmpOld = SelectBitmap(hdcMem, this->bitmap);
                BLENDFUNCTION bf = { AC_SRC_OVER, 0, 255, AC_SRC_ALPHA };
                ::GdiAlphaBlend(hdc, x, y, width, height, hdcMem, 0, 0, this->width, this->height, bf);
                SelectBitmap(hdcMem, bmpOld);
                ::DeleteDC(hdcMem);
            }
            ::EndPaint(hwnd, &ps);
        }
    }

    double ScaleSize(int width, int height, int* newWidth, int* newHeight) const throw()
    {
        const double xScale = this->width <= width ? 1 : static_cast<double>(width) / this->width;
        const double yScale = this->height <= height ? 1 : static_cast<double>(height) / this->height;
        const double scale = xScale < yScale ? xScale : yScale;
        *newWidth  = (int)floor(scale * this->width);
        *newHeight = (int)floor(scale * this->height);
        return scale;
    }

    DROPEFFECT OnDragEnter(IDataObject*, DWORD, POINT) throw()
    {
        return DROPEFFECT_COPY;
    }

    DROPEFFECT OnDragOver(DWORD, POINT) throw()
    {
        return DROPEFFECT_COPY;
    }

    void OnDragLeave() throw()
    {
    }

    DROPEFFECT OnDrop(IDataObject* dataObject, DWORD, POINT) throw()
    {
        this->dataObject = dataObject;
        if (this->bitmap != NULL) {
            DeleteBitmap(this->bitmap);
            this->bitmap = NULL;
        }
        ComPtr<IShellItemImageFactory> factory;
        HRESULT hr = ::SHGetItemFromDataObject(dataObject, DOGIF_DEFAULT, IID_PPV_ARGS(&factory));
        if (SUCCEEDED(hr)) {
            const SIZE size = { 256, 256 };
            hr = factory->GetImage(size, SIIGBF_RESIZETOFIT, &this->bitmap);
        }
        if (SUCCEEDED(hr)) {
            BITMAP bm = {};
            ::GetObject(this->bitmap, sizeof(bm), &bm);
            this->width = bm.bmWidth;
            this->height = bm.bmHeight;
        }
        HWND hwnd = NULL;
        if (SUCCEEDED(this->GetWindow(&hwnd))) {
            ::InvalidateRect(hwnd, NULL, FALSE);
        }
        return DROPEFFECT_COPY;
    }

public:
    MainWindow() throw()
        : bitmap()
        , width()
        , height()
    {
    }

    ~MainWindow() throw()
    {
        if (this->bitmap != NULL) {
            DeleteBitmap(this->bitmap);
        }
    }

    HWND Create() throw()
    {
        ::InitCommonControls();
        const ATOM atom = __super::Register(MAKEINTATOM(411));
        if (atom == 0) {
            return NULL;
        }
        return __super::Create(atom, L"Drag Drop", WS_OVERLAPPEDWINDOW, WS_EX_TOOLWINDOW);
    }
};


class COleInitialize
{
public:
    HRESULT m_hr;

    COleInitialize() throw()
        : m_hr(::OleInitialize(NULL))
    {
    }

    ~COleInitialize() throw()
    {
        if (SUCCEEDED(m_hr)) {
            ::OleUninitialize();
        }
    }

    HRESULT GetResult() const throw()
    {
        return m_hr;
    }
};

class CAllocConsole
{
public:
    CAllocConsole() throw()
    {
        ::AllocConsole();
    }
    ~CAllocConsole() throw()
    {
        ::FreeConsole();
    }
};

int WINAPI wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int) throw()
{
#ifdef USETRACE
    CAllocConsole con;
#endif
    COleInitialize init;
    MainWindow wnd;
    HWND hwnd = wnd.Create();
    if (hwnd == NULL) {
        return EXIT_FAILURE;
    }
    RECT rect = { 0, 0, 256, 256 };
    MyAdjustWindowRect(hwnd, &rect);
    ::SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, rect.right - rect.left, rect.bottom - rect.top, SWP_NOMOVE | SWP_NOACTIVATE);
    ::ShowWindow(hwnd, SW_SHOW);
    ::UpdateWindow(hwnd);
    MSG msg;
    msg.wParam = EXIT_FAILURE;
    while (::GetMessage(&msg, NULL, 0, 0)) {
        ::DispatchMessage(&msg);
        ::TranslateMessage(&msg);
    }
    return static_cast<int>(msg.wParam);
}

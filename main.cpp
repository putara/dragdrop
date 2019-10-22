#define WINVER          0x0601
#define _WIN32_WINNT    0x0601
#define _WIN32_IE       0x0800
#define UNICODE
#define _UNICODE
#define WINBASE_DECLARE_RESTORE_LAST_ERROR

#include <sdkddkver.h>
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <shlwapi.h>
#include <thumbcache.h>
#include <wincodec.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <vsstyle.h>
#include <vssym32.h>
#include <stddef.h>
#include <intrin.h>
#include <math.h>

#include <strsafe.h>

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

#define HANDLE_WM_DWMCOMPOSITIONCHANGED(hwnd, wParam, lParam, fn) \
    ((fn)((hwnd)), 0L)


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
    void Attach(T* src) throw()
    {
        this->Release();
        this->ptr = src;
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


template <class T>
class TPtrArrayDestroyHelper
{
public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_Destroy(hdpa);
    }
};


template <class T>
class TPtrArrayAutoDeleteHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        delete static_cast<T*>(p);
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};


template <class T>
class TPtrArrayAutoReleaseHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        static_cast<T*>(p)->Release();
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};


template <class T>
class TPtrArrayAutoFreeHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        ::free(p);
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};


template <class T>
class TPtrArrayAutoCoFreeHelper
{
    static int CALLBACK sDeletePtrCB(void* p, void*)
    {
        ::SHFree(p);
        return 1;
    }

public:
    static void DeleteAllPtrs(HDPA hdpa)
    {
        ::DPA_EnumCallback(hdpa, sDeletePtrCB, NULL);
        ::DPA_DeleteAllPtrs(hdpa);
    }
    static void Destroy(HDPA hdpa)
    {
        ::DPA_DestroyCallback(hdpa, sDeletePtrCB, NULL);
    }
};


template <class T, int cGrow, typename TDestroy>
class PtrArray
{
private:
    HDPA hdpa;

    PtrArray(const PtrArray&) throw();
    PtrArray& operator =(const PtrArray&) throw();

public:
    PtrArray()
        : hdpa()
    {
    }
    ~PtrArray()
    {
        TDestroy::Destroy(this->hdpa);
    }

    int GetCount() const throw()
    {
        if (this->hdpa == NULL) {
            return 0;
        }
        return DPA_GetPtrCount(this->hdpa);
    }

    T* FastGetAt(int index) const throw()
    {
        return static_cast<T*>(DPA_FastGetPtr(this->hdpa, index));
    }

    T* GetAt(int index) const throw()
    {
        return static_cast<T*>(::DPA_GetPtr(this->hdpa, index));
    }

    T** GetData() const throw()
    {
        if (this->hdpa == NULL) {
            return NULL;
        }
        return reinterpret_cast<T**>(DPA_GetPtrPtr(this->hdpa));
    }

    bool Grow(int newSize) throw()
    {
        return (::DPA_Grow(this->hdpa, newSize) != FALSE);
    }

    bool SetAtGrow(int index, T* p) throw()
    {
        return (::DPA_SetPtr(this->hdpa, index, p) != FALSE);
    }

    bool Create() throw()
    {
        if (this->hdpa != NULL) {
            return true;
        }
        this->hdpa = ::DPA_Create(cGrow);
        return (this->hdpa != NULL);
    }

    bool Create(HANDLE hheap) throw()
    {
        if (this->hdpa != NULL) {
            return true;
        }
        this->hdpa = ::DPA_CreateEx(cGrow, hheap);
        return (this->hdpa != NULL);
    }

    int Add(T* p) throw()
    {
        return InsertAt(DA_LAST, p);
    }

    int InsertAt(int index, T* p) throw()
    {
        return ::DPA_InsertPtr(this->hdpa, index, p);
    }

    T* RemoveAt(int index) throw()
    {
        return static_cast<T*>(::DPA_DeletePtr(this->hdpa, index));
    }

    void RemoveAll()
    {
        TDestroy::DeleteAllPtrs(this->hdpa);
    }

    void Enumerate(__in PFNDAENUMCALLBACK callback, __in_opt void* context) const
    {
        ::DPA_EnumCallback(this->hdpa, callback, context);
    }

    int Search(__in_opt void* key, __in int start, __in PFNDACOMPARE compare, __in LPARAM lParam, __in UINT options)
    {
        return ::DPA_Search(this->hdpa, key, start, compare, lParam, options);
    }

    bool Sort(__in PFNDACOMPARE compare, __in LPARAM lParam)
    {
        return (::DPA_Sort(this->hdpa, compare, lParam) != FALSE);
    }

    T* operator [](int index) const throw()
    {
        return GetAt(index);
    }
};


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


class DataObject : public IDataObject, public IPersist
{
protected:
    struct Item
    {
        FORMATETC fetc;
        STGMEDIUM medium;

        Item() throw()
        {
            this->Clear();
        }

        ~Item() throw()
        {
            this->Release();
        }

        void Clear() throw()
        {
            ZeroMemory(&this->fetc, sizeof(this->fetc));
            ZeroMemory(&this->medium, sizeof(this->medium));
        }

        void Release() throw()
        {
            ::CoTaskMemFree(this->fetc.ptd);
            ::ReleaseStgMedium(&this->medium);
        }
    };

    typedef PtrArray<Item, 0, TPtrArrayAutoDeleteHelper<Item>> ItemArray;

    LONG volatile cRef;
    ItemArray items;

    static const CLSID CLSID_MyDataObject;

public:
    DataObject() throw()
        : cRef(1)
    {
    }

    ~DataObject() throw()
    {
    }

    STDMETHOD(QueryInterface)(REFIID iid, __deref_out void** ppv) throw()
    {
        if (IsEqualIID(iid, __uuidof(IUnknown)) || IsEqualIID(iid, __uuidof(IDataObject))) {
            *ppv = static_cast<IDataObject*>(this);
            this->AddRef();
            return S_OK;
        }
        if (IsEqualIID(iid, __uuidof(IPersist))) {
            *ppv = static_cast<IPersist*>(this);
            this->AddRef();
            return S_OK;
        }

        *ppv = NULL;
        return E_NOINTERFACE;
    }

    STDMETHOD_(ULONG, AddRef)() throw()
    {
        return ::InterlockedIncrement(&this->cRef);
    }

    STDMETHOD_(ULONG, Release)() throw()
    {
        LONG ret = ::InterlockedDecrement(&this->cRef);
        if (ret == 0) {
            delete this;
        }
        return ret;
    }

    STDMETHOD(GetData)(FORMATETC* fetc, STGMEDIUM* medium) throw()
    {
        TRACE("IDataObject::GetData\n");

        if (fetc == NULL || medium == NULL) {
            return E_INVALIDARG;
        }

        TRACE("  cfFormat = %u\n", fetc->cfFormat);
        TRACE("  TYMED = %x\n", fetc->tymed);

        Item* item;
        HRESULT hr = this->FindFORMATETC(fetc, false, &item);
        if (SUCCEEDED(hr)) {
            hr = this->AddRefStgMedium(&item->medium, medium, true);
        }
        return hr;
    }

    STDMETHOD(QueryGetData)(FORMATETC* fetc) throw()
    {
        TRACE("IDataObject::QueryGetData\n");

        if (fetc == NULL) {
            return E_INVALIDARG;
        }

        TRACE("  cfFormat = %u\n", fetc->cfFormat);
        TRACE("  TYMED = %x\n", fetc->tymed);

        Item* item;
        return this->FindFORMATETC(fetc, false, &item);
    }

    STDMETHOD(GetDataHere)(FORMATETC*, STGMEDIUM*) throw()
    {
        TRACE("IDataObject::GetDataHere - E_NOTIMPL\n");
        return E_NOTIMPL;
    }

    STDMETHOD(GetCanonicalFormatEtc)(FORMATETC*, FORMATETC*) throw()
    {
        TRACE("IDataObject::GetCanonicalFormatEtc - E_NOTIMPL\n");
        return E_NOTIMPL;
    }

    STDMETHOD(SetData)(FORMATETC* fetc, STGMEDIUM* medium, BOOL fRelease) throw()
    {
        TRACE("IDataObject::SetData\n");

        if (fetc == NULL || medium == NULL) {
            return E_INVALIDARG;
        }

        Item* item;
        HRESULT hr = this->FindFORMATETC(fetc, true, &item);
        if (SUCCEEDED(hr)) {
            if (item->medium.tymed) {
                ::ReleaseStgMedium(&item->medium);
            }

            if (fRelease) {
                item->medium = *medium;
                hr = S_OK;
            } else {
                hr = this->AddRefStgMedium(medium, &item->medium, true);
            }
            item->fetc.tymed = item->medium.tymed;

            IUnknown* unkMedium = GetCanonicalIUnknown(item->medium.pUnkForRelease);
            IUnknown* unkSelf   = GetCanonicalIUnknown(static_cast<IDataObject*>(this));
            if (unkMedium == unkSelf) {
                item->medium.pUnkForRelease->Release();
                item->medium.pUnkForRelease = NULL;
            }
        }
        return hr;
    }

    STDMETHOD(EnumFormatEtc)(DWORD direction, __deref_out IEnumFORMATETC** ppenumFetc) throw()
    {
        TRACE("IDataObject::EnumFormatEtc\n");

        if (ppenumFetc == NULL) {
            return E_POINTER;
        }
        *ppenumFetc = NULL;

        if (direction != DATADIR_GET) {
            return E_NOTIMPL;
        }

        Item** items = this->items.GetData();
        const int cItems = this->items.GetCount();
        FORMATETC* fetc = new FORMATETC[cItems];
        if (fetc == NULL) {
            return E_OUTOFMEMORY;
        }
        for (int i = 0; i < cItems; i++) {
            fetc[i] = items[i]->fetc;
        }
        HRESULT hr = ::SHCreateStdEnumFmtEtc(cItems, fetc, ppenumFetc);
        delete[] fetc;
        return hr;
    }

    STDMETHOD(DAdvise)(FORMATETC*, DWORD, IAdviseSink*, DWORD*) throw()
    {
        TRACE("IDataObject::DAdvise - E_NOTIMPL\n");
        return E_NOTIMPL;
    }

    STDMETHOD(DUnadvise)(DWORD) throw()
    {
        TRACE("IDataObject::DUnadvise - E_NOTIMPL\n");
        return E_NOTIMPL;
    }

    STDMETHOD(EnumDAdvise)(IEnumSTATDATA** ppenumAdvise) throw()
    {
        TRACE("IDataObject::EnumDAdvise - E_NOTIMPL\n");

        if (ppenumAdvise == NULL) {
            return E_POINTER;
        }
        *ppenumAdvise = NULL;
        return E_NOTIMPL;
    }

    STDMETHOD(GetClassID)(__out CLSID* clsid) throw()
    {
        *clsid = CLSID_MyDataObject;
        return S_OK;
    }

    static HRESULT Clone(IDataObject* dataObjectSrc, __deref_out IDataObject** dataObjectDst) throw()
    {
        *dataObjectDst = NULL;
        ComPtr<DataObject> self;
        ComPtr<IEnumFORMATETC> enumFetc;
        self.Attach(new DataObject());
        HRESULT hr = self != NULL ? S_OK : E_OUTOFMEMORY;
        if (SUCCEEDED(hr)) {
            hr = dataObjectSrc->EnumFormatEtc(DATADIR_GET, &enumFetc);
        }
        if (SUCCEEDED(hr)) {
            enumFetc->Reset();
            for (Item item; SUCCEEDED(hr); item.Clear()) {
                hr = enumFetc->Next(1, &item.fetc, NULL);
                if (SUCCEEDED(hr)) {
                    if (hr != S_OK) {
                        hr = S_OK;
                        break;
                    }
                    hr = dataObjectSrc->GetData(&item.fetc, &item.medium);
                    if (FAILED(hr)) {
                        item.Release();
                        // ignore failure
                        hr = S_OK;
                        continue;
                    }
                }
                if (SUCCEEDED(hr)) {
                    hr = self->SetData(&item.fetc, &item.medium, TRUE);
                }
                if (FAILED(hr)) {
                    break;
                }
            }
        }
        if (SUCCEEDED(hr)) {
            *dataObjectDst = self.Detach();
        }
        return hr;
    }

    static BOOL IsMyDataObject(IDataObject* dataObject) throw()
    {
        ComPtr<IPersist> persist;
        if (SUCCEEDED(dataObject->QueryInterface(IID_PPV_ARGS(&persist)))) {
            CLSID clsid = {};
            if (SUCCEEDED(persist->GetClassID(&clsid))) {
                return IsEqualCLSID(clsid, CLSID_MyDataObject);
            }
        }
        return FALSE;
    }

protected:
    HRESULT FindFORMATETC(FORMATETC* fetc, bool addIfNotFound, __deref_out Item** found) throw()
    {
        *found = NULL;
        if ((fetc->dwAspect & DVASPECT_CONTENT) == 0) {
            TRACE("  - DV_E_DVASPECT\n");
            return DV_E_DVASPECT;
        }

        if (fetc->ptd != NULL) {
            TRACE("  - DV_E_DVTARGETDEVICE\n");
            return DV_E_DVTARGETDEVICE;
        }

        Item** items = this->items.GetData();
        const int cItems = this->items.GetCount();
        for (int i = 0; i < cItems; i++) {
            if (items[i]->fetc.cfFormat != fetc->cfFormat
                    || (items[i]->fetc.tymed & fetc->tymed) == 0
                    || (items[i]->fetc.lindex == -1 && items[i]->fetc.lindex != fetc->lindex)) {
                continue;
            }

            if (addIfNotFound || (items[i]->fetc.tymed & fetc->tymed)) {
                TRACE("  - S_OK\n");
                *found = items[i];
                return S_OK;
            }

            TRACE("  - DV_E_TYMED\n");
            return DV_E_TYMED;
        }

        if (addIfNotFound == false) {
            TRACE("  - DATA_E_FORMATETC\n");
            return DATA_E_FORMATETC;
        }

        Item* item = new Item();
        if (item == NULL) {
            return E_OUTOFMEMORY;
        }
        CopyMemory(&item->fetc, fetc, sizeof(FORMATETC));

        this->items.Create();
        if (this->items.Add(item) == -1) {
            TRACE("  - E_OUTOFMEMORY\n");
            delete item;
            return E_OUTOFMEMORY;
        }

        *found = item;
        TRACE("  - S_OK\n");
        return S_OK;
    }

    HRESULT AddRefStgMedium(__in STGMEDIUM* mediumIn, __out STGMEDIUM* mediumOut, bool copyIn) throw()
    {
        HRESULT hr = S_OK;
        STGMEDIUM temp = *mediumIn;

        if (mediumIn->pUnkForRelease == NULL && (mediumIn->tymed & (TYMED_ISTREAM | TYMED_ISTORAGE)) == 0) {
            if (copyIn) {
                if (mediumIn->tymed != TYMED_HGLOBAL) {
                    return DV_E_TYMED;
                }

                temp.hGlobal = GlobalClone(mediumIn->hGlobal);
                if (temp.hGlobal == NULL) {
                    return E_OUTOFMEMORY;
                }
            } else {
                temp.pUnkForRelease = static_cast<IDataObject*>(this);
            }
        }

        if (SUCCEEDED(hr)) {
            switch (temp.tymed) {
            case TYMED_ISTREAM:
                temp.pstm->AddRef();
                break;
            case TYMED_ISTORAGE:
                temp.pstg->AddRef();
                break;
            }

            if (temp.pUnkForRelease != NULL) {
                temp.pUnkForRelease->AddRef();
            }
            *mediumOut = temp;
        }
        return hr;
    }

    static HGLOBAL GlobalClone(HGLOBAL hglobSrc) throw()
    {
        if (hglobSrc != NULL) {
            const size_t cbSize = ::GlobalSize(hglobSrc);
            void* const srcPtr = ::GlobalLock(hglobSrc);
            if (srcPtr != NULL) {
                HGLOBAL hglobDst = ::GlobalAlloc(GMEM_MOVEABLE, cbSize);
                if (hglobDst != NULL) {
                    void* const dstPtr = ::GlobalLock(hglobDst);
                    if (dstPtr != NULL) {
                        CopyMemory(dstPtr, srcPtr, cbSize);
                        ::GlobalUnlock(hglobDst);
                        ::GlobalUnlock(hglobSrc);
                        return hglobDst;
                    }
                    ::GlobalFree(hglobDst);
                }
                ::GlobalUnlock(hglobSrc);
            }
        }
        return NULL;
    }

    static IUnknown* GetCanonicalIUnknown(IUnknown* unk) throw()
    {
        IUnknown* unkCanonical = NULL;
        if (unk != NULL && SUCCEEDED(unk->QueryInterface(IID_PPV_ARGS(&unkCanonical)))) {
            unkCanonical->Release();
        } else {
            unkCanonical = unk;
        }
        return unkCanonical;
    }
};

// {6DB994BF-6A3C-4355-BCEF-9A412221E033}
const CLSID DataObject::CLSID_MyDataObject = { 0x6db994bf, 0x6a3c, 0x4355, { 0xbc, 0xef, 0x9a, 0x41, 0x22, 0x21, 0xe0, 0x33 } };


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

    STDMETHOD(QueryInterface)(REFIID iid, __deref_out void** ppv) throw()
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


#include "res.rc"

class MainWindow : public BaseWindow<MainWindow>, public ResizableWindow<MainWindow>, public DropTargetExImpl<MainWindow>
{
protected:
    friend BaseWindow<MainWindow>;
    friend DropTargetImpl<MainWindow>;
    friend DropTargetExImpl<MainWindow>;

    HWND toolbar;
    HIMAGELIST closeImage;
    HBITMAP bitmap;
    HBITMAP bitmapPM;
    int width;
    int height;
    HFONT font;
    int textHeight;
    LPWSTR name;
    ComPtr<IDataObject> dataObject;
    MagneticWindow mag;

    static const int BTN_SIZE = 16;

    static void ErrorMessage(HWND hwnd, LPCWSTR name) throw()
    {
        WCHAR message[MAX_PATH + 64];
        ::StringCchCopyW(message, _countof(message), L"Could not copy:\n");
        ::StringCchCatW(message, _countof(message), name);
        ::MessageBoxW(hwnd, message, L"Sorry", MB_OK);
    }

    LRESULT CALLBACK WindowProc(HWND hwnd, UINT message, WPARAM wParam, LPARAM lParam) throw()
    {
        LRESULT result;
        if (ResizableWindow<MainWindow>::ProcessWindowMessages(hwnd, message, wParam, lParam, &result)) {
            return result;
        }

        switch (message) {
        HANDLE_MSG(hwnd, WM_CREATE, OnCreate);
        HANDLE_MSG(hwnd, WM_DESTROY, OnDestroy);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
        HANDLE_MSG(hwnd, WM_WINDOWPOSCHANGED, OnWindowPosChanged);
        HANDLE_MSG(hwnd, WM_SETTINGCHANGE, OnSettingChange);
        HANDLE_MSG(hwnd, WM_ACTIVATE, OnActivate);
        HANDLE_MSG(hwnd, WM_NCCALCSIZE, OnNcCalcSize);
        HANDLE_MSG(hwnd, WM_NOTIFY, OnNotify);
        HANDLE_MSG(hwnd, WM_INITMENU, OnInitMenu);
        HANDLE_MSG(hwnd, WM_INITMENUPOPUP, OnInitMenuPopup);
        HANDLE_MSG(hwnd, WM_CONTEXTMENU, OnContextMenu);
        HANDLE_MSG(hwnd, WM_LBUTTONDOWN, OnButtonDown);
        HANDLE_MSG(hwnd, WM_RBUTTONDOWN, OnButtonDown);
        HANDLE_MSG(hwnd, WM_MBUTTONDOWN, OnButtonDown);
        HANDLE_MSG(hwnd, WM_MOUSEMOVE, OnMouseMove);
        HANDLE_MSG(hwnd, WM_LBUTTONUP, OnLButtonUp);
        HANDLE_MSG(hwnd, WM_PAINT, OnPaint);
        HANDLE_MSG(hwnd, WM_DWMCOMPOSITIONCHANGED, OnDwmCompositionChanged);
        }
        return ::DefWindowProc(hwnd, message, wParam, lParam);
    }

    bool OnCreate(HWND hwnd, LPCREATESTRUCT lpcs) throw()
    {
        if (lpcs == NULL) {
            return false;
        }
        this->RegisterDragDrop(hwnd);
        this->InitFont(hwnd);
        this->InitCloseButton(hwnd);
        this->InitDwmAttrs(hwnd);
        HMENU menu = ::GetSystemMenu(hwnd, FALSE);
        ::RemoveMenu(menu, SC_RESTORE, MF_BYCOMMAND);
        ::RemoveMenu(menu, SC_MINIMIZE, MF_BYCOMMAND);
        ::RemoveMenu(menu, SC_MAXIMIZE, MF_BYCOMMAND);
        return true;
    }

    void OnDestroy(HWND) throw()
    {
        this->dataObject.Release();
        this->RevokeDragDrop();
        ::PostQuitMessage(0);
    }

    void OnCommand(HWND hwnd, int id, HWND, UINT) throw()
    {
        if (id == SC_CLOSE) {
            ::PostMessage(hwnd, WM_CLOSE, 0, 0);
        }
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
            if (this->toolbar != NULL) {
                RECT rect = {}, btnRect = {};
                ::GetClientRect(hwnd, &rect);
                ::SendMessage(this->toolbar, TB_GETITEMRECT, 0, reinterpret_cast<LPARAM>(&btnRect));
                ::SetWindowPos(this->toolbar, NULL, rect.right - btnRect.right - 4, 4, btnRect.right, btnRect.bottom, 0);
            }
            ::InvalidateRect(hwnd, NULL, FALSE);
        }
    }

    void OnSettingChange(HWND hwnd, UINT, LPCTSTR) throw()
    {
        this->InitFont(hwnd);
        this->InitDwmAttrs(hwnd);
        ::InvalidateRect(hwnd, NULL, FALSE);
    }

    void OnActivate(HWND hwnd, UINT, HWND, BOOL) throw()
    {
        ::SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
    }

    UINT OnNcCalcSize(HWND, BOOL, NCCALCSIZE_PARAMS*) throw()
    {
        return 0;
    }

    LRESULT OnNotify(HWND, int, NMHDR* pnmh) throw()
    {
        if (pnmh->code == NM_CUSTOMDRAW && this->toolbar != NULL && pnmh->hwndFrom == this->toolbar) {
            LPNMTBCUSTOMDRAW tbcd = CONTAINING_RECORD(pnmh, NMTBCUSTOMDRAW, nmcd.hdr);
            switch (tbcd->nmcd.dwDrawStage) {
            case CDDS_PREPAINT:
                return this->OnNotifyPrePaintTBCD(tbcd);

            case CDDS_ITEMPREPAINT:
                return this->OnNotifyItemPrePaintTBCD(tbcd);
            }
            return 0;
        }
        return 0;
    }

    LRESULT OnNotifyPrePaintTBCD(LPNMTBCUSTOMDRAW tbcd) throw()
    {
        RECT rect = tbcd->nmcd.rc;
        ::PatBlt(tbcd->nmcd.hdc, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, BLACKNESS);
        return CDRF_NOTIFYITEMDRAW;
    }

    LRESULT OnNotifyItemPrePaintTBCD(LPNMTBCUSTOMDRAW tbcd) throw()
    {
        this->OnNotifyPrePaintTBCD(tbcd);
        const UINT state = tbcd->nmcd.uItemState;
        int image = -1;
        if (state & CDIS_SELECTED) {
            image = 1;
        } else if (state & CDIS_HOT) {
            image = 0;
        }
        if (image >= 0) {
            ::ImageList_Draw(this->closeImage, image, tbcd->nmcd.hdc, tbcd->nmcd.rc.left, tbcd->nmcd.rc.top, ILD_NORMAL);
        }
        return CDRF_SKIPDEFAULT;
    }

    void OnInitMenu(HWND, HMENU) throw()
    {
    }

    void OnInitMenuPopup(HWND, HMENU, UINT, BOOL) throw()
    {
    }

    void OnContextMenu(HWND hwnd, HWND hwndContext, int xPos, int yPos) throw()
    {
        ::PostMessage(hwnd, 0x313, 0, MAKELPARAM(xPos, yPos));
        FORWARD_WM_CONTEXTMENU(hwnd, hwndContext, xPos, yPos, ::DefWindowProc);
    }

    void OnButtonDown(HWND hwnd, BOOL, int x, int y, UINT flags) throw()
    {
        if (this->dataObject != NULL) {
            POINT pt = {};
            ::GetCursorPos(&pt);
            RECT rect = {};
            ::GetClientRect(hwnd, &rect);
            SIZE size = { rect.right, rect.bottom };
            ::InflateRect(&rect, -16, -16);
            MapWindowRect(hwnd, HWND_DESKTOP, &rect);
            if (::PtInRect(&rect, pt)) {
                if (MyDragDetect(hwnd, pt)) {
                    ComPtr<IDragSourceHelper2> helper;
                    if (SUCCEEDED(::CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&helper)))) {
                        int left, top, width, height;
                        GetBitmapArea(size, &left, &top, &width, &height);
                        SHDRAGIMAGE shdi = {
                            { this->width, this->height },
                            {
                                ::MulDiv(x - left, this->width, width),
                                ::MulDiv(y - top, this->height, height)
                            },
                            static_cast<HBITMAP>(::CopyImage(this->bitmap, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)),
                            0
                        };
                        if (FAILED(helper->InitializeFromBitmap(&shdi, this->dataObject))) {
                            DeleteBitmap(shdi.hbmpDragImage);
                        }
                        helper->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
                    }
                    DWORD effect = DROPEFFECT_COPY;
                    ::SHDoDragDrop(hwnd, this->dataObject, NULL, effect, &effect);
                }
                return;
            }
        }
        if (flags & MK_LBUTTON) {
            this->mag.BeginDrag(hwnd, x, y);
        }
    }

    void OnMouseMove(HWND hwnd, int x, int y, UINT) throw()
    {
        this->mag.Drag(hwnd, x, y);
    }

    void OnLButtonUp(HWND, int, int, UINT) throw()
    {
        this->mag.EndDrag();
    }

    void OnPaint(HWND hwnd) throw()
    {
        PAINTSTRUCT ps = {};
        if (::BeginPaint(hwnd, &ps) != NULL) {
            RECT rect = {};
            ::GetClientRect(hwnd, &rect);
            const SIZE size = { rect.right, rect.bottom };
            BP_PAINTPARAMS bppp = { sizeof(bppp) };
            bppp.dwFlags = BPPF_ERASE;
            HDC hdc = NULL;
            HPAINTBUFFER hpb = ::BeginBufferedPaint(ps.hdc, &rect, BPBF_COMPOSITED, &bppp, &hdc);
            if (hpb != NULL) {
                this->PaintContent(hwnd, hdc, size);
                ::EndBufferedPaint(hpb, TRUE);
            }
            ::EndPaint(hwnd, &ps);
        }
    }

    void OnDwmCompositionChanged(HWND hwnd) throw()
    {
        this->InitDwmAttrs(hwnd);
    }

    void PaintContent(HWND hwnd, HDC hdc, const SIZE size) throw()
    {
        this->DrawBitmap(hwnd, hdc, size);
        this->DrawText(hwnd, hdc, size);
    }

    void GetBitmapArea(__in SIZE size, __out int* x, __out int* y, __out int* width, __out int* height) throw()
    {
        if (size.cy >= this->textHeight) {
            size.cy -= this->textHeight;
        }
        double scale = this->ScaleSize(size.cx, size.cy, width, height);
        scale; // unused
        *x = (size.cx - *width) / 2;
        *y = (size.cy - *height) / 2;
    }

    void DrawBitmap(HWND, HDC hdc, SIZE size) throw()
    {
        if (this->bitmapPM != NULL) {
            int x, y, width, height;
            GetBitmapArea(size, &x, &y, &width, &height);
            HDC hdcMem = ::CreateCompatibleDC(hdc);
            HBITMAP bmpOld = SelectBitmap(hdcMem, this->bitmapPM);
            BLENDFUNCTION bf = { AC_SRC_OVER, 0, 255, AC_SRC_ALPHA };
            int mode = ::SetStretchBltMode(hdc, HALFTONE);
            ::GdiAlphaBlend(hdc, x, y, width, height, hdcMem, 0, 0, this->width, this->height, bf);
            ::SetStretchBltMode(hdc, mode);
            SelectBitmap(hdcMem, bmpOld);
            ::DeleteDC(hdcMem);
        }
    }

    void DrawText(HWND hwnd, HDC hdc, const SIZE size) throw()
    {
        if (this->name != NULL) {
            if (size.cy > this->textHeight) {
                HFONT fontOld = SelectFont(hdc, this->font);
                RECT rect = { 0, size.cy - this->textHeight, size.cx, size.cy };
                if (rect.top < 0) {
                    rect.top = 0;
                }
                DTTOPTS dtto = { sizeof(dtto) };
                dtto.dwFlags = DTT_TEXTCOLOR | DTT_COMPOSITED;
                dtto.crText = 0;
                HTHEME theme = ::OpenThemeData(hwnd, L"globals");
                for (int y = -1; y <= 1; y++) {
                    for (int x = -1; x <= 1; x++) {
                        if (x != 0 && y != 0) {
                            RECT shadowRect = rect;
                            ::OffsetRect(&shadowRect, x, y);
                            ::DrawThemeTextEx(theme, hdc, TEXT_BODYTITLE, 0, this->name, -1, DT_CENTER | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS, &shadowRect, &dtto);
                        }
                    }
                }
                dtto.crText = 0xffffff;
                ::DrawThemeTextEx(theme, hdc, TEXT_BODYTITLE, 0, this->name, -1, DT_CENTER | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS, &rect, &dtto);
                ::CloseThemeData(theme);
                SelectFont(hdc, fontOld);
            }
        }
    }

    void InitFont(HWND hwnd) throw()
    {
        if (this->font != NULL) {
            DeleteFont(this->font);
        }
        NONCLIENTMETRICS ncm = { sizeof(ncm) };
        ::SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
        // ncm.lfStatusFont.lfHeight = ncm.lfStatusFont.lfHeight * 4 / 3;
        // ncm.lfStatusFont.lfWidth = 0;
        ncm.lfStatusFont.lfWeight += FW_BOLD;
        ncm.lfStatusFont.lfQuality = ANTIALIASED_QUALITY;
        this->font = ::CreateFontIndirect(&ncm.lfStatusFont);
        HDC hdc = ::GetDC(hwnd);
        HFONT fontOld = SelectFont(hdc, this->font);
        TEXTMETRIC tm = {};
        ::GetTextMetrics(hdc, &tm);
        this->textHeight = tm.tmHeight + 8;
        SelectFont(hdc, fontOld);
        ::ReleaseDC(hwnd, hdc);
    }

    void InitCloseButton(HWND hwnd) throw()
    {
        const DWORD WS_TOOLBAR_DEFAULT
            = WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN
            | CCS_TOP | CCS_NODIVIDER | CCS_NORESIZE | CCS_NOPARENTALIGN
            | TBSTYLE_TOOLTIPS | TBSTYLE_FLAT | TBSTYLE_TRANSPARENT;
        const TBBUTTON btn = { 0, SC_CLOSE, TBSTATE_ENABLED, BTNS_BUTTON, 0, 0 };
        HWND toolbar = ::CreateToolbarEx(
                            hwnd
                            , WS_TOOLBAR_DEFAULT
                            , 1, 0, NULL, 0, &btn, 1
                            , BTN_SIZE, BTN_SIZE, 0, 0, sizeof(TBBUTTON));
        if (toolbar == NULL) {
            return;
        }
        ::SendMessage(toolbar, TB_SETPARENT, reinterpret_cast<WPARAM>(hwnd), 0);
        ::SendMessage(toolbar, CCM_SETVERSION, 6, 0);
        ::SendMessage(toolbar, TB_SETEXTENDEDSTYLE, 0, TBSTYLE_EX_DOUBLEBUFFER);
        TBMETRICS tbm = { sizeof(tbm), TBMF_PAD | TBMF_BUTTONSPACING };
        ::SendMessage(toolbar, TB_SETMETRICS, 0, reinterpret_cast<LPARAM>(&tbm));
        ::SendMessage(toolbar, TB_SETBITMAPSIZE, 0, MAKELPARAM(BTN_SIZE, BTN_SIZE));
        ::SendMessage(toolbar, TB_SETBUTTONSIZE, 0, MAKELPARAM(BTN_SIZE, BTN_SIZE));
        ::SendMessage(toolbar, TB_AUTOSIZE, 0, 0);
        this->toolbar = toolbar;
        this->closeImage = ::ImageList_LoadImage(HINST_THISCOMPONENT, MAKEINTRESOURCE(IDB_CLOSE), BTN_SIZE, 0, 0, IMAGE_BITMAP, LR_CREATEDIBSECTION);
    }

    static HRESULT DwmSetWindowAttributeDword(HWND hwnd, DWMWINDOWATTRIBUTE attribute, DWORD value) throw()
    {
        return ::DwmSetWindowAttribute(hwnd, attribute, &value, sizeof(value));
    }

    void InitDwmAttrs(HWND hwnd) throw()
    {
        static const MARGINS margins = { -1, -1, -1, -1 };
        ::DwmExtendFrameIntoClientArea(hwnd, &margins);
        DwmSetWindowAttributeDword(hwnd, DWMWA_ALLOW_NCPAINT, true);
        DwmSetWindowAttributeDword(hwnd, DWMWA_NCRENDERING_POLICY, DWMNCRP_DISABLED);
        ::SetWindowBlur(hwnd);
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

    DROPEFFECT OnDragEnter(IDataObject* dataObject, DWORD, POINT) throw()
    {
        if (DataObject::IsMyDataObject(dataObject)) {
            return DROPEFFECT_NONE;
        }
        return DROPEFFECT_COPY;
    }

    DROPEFFECT OnDragOver(DWORD, POINT) throw()
    {
        if (DataObject::IsMyDataObject(this->dataObjectDragging)) {
            return DROPEFFECT_NONE;
        }
        return DROPEFFECT_COPY;
    }

    void OnDragLeave() throw()
    {
    }

    DROPEFFECT OnDrop(IDataObject* dataObject, DWORD, POINT) throw()
    {
        if (this->SetObject(dataObject)) {
            return DROPEFFECT_COPY;
        }
        return DROPEFFECT_NONE;
    }

    static HRESULT GetBitmap(__in IDataObject* dataObject, __out HBITMAP* bmp) throw()
    {
        const int IMAGE_SIZE = 256;
        *bmp = NULL;
        ComPtr<IShellItem> item;
        HRESULT hr = ::SHGetItemFromDataObject(dataObject, DOGIF_DEFAULT, IID_PPV_ARGS(&item));
        if (FAILED(hr)) {
            return hr;
        }
        ComPtr<IThumbnailProvider> thump;
        hr = item->BindToHandler(NULL, BHID_ThumbnailHandler, IID_PPV_ARGS(&thump));
        if (SUCCEEDED(hr)) {
            WTS_ALPHATYPE alpha = WTSAT_UNKNOWN;
            hr = thump->GetThumbnail(IMAGE_SIZE, bmp, &alpha);
            if (SUCCEEDED(hr)) {
                if (alpha == WTSAT_ARGB) {
                    return hr;
                }
                DeleteBitmap(*bmp);
                *bmp = NULL;
            }
        }
        ComPtr<IShellItemImageFactory> factory;
        hr = item->QueryInterface(IID_PPV_ARGS(&factory));
        if (SUCCEEDED(hr)) {
            const SIZE size = { IMAGE_SIZE, IMAGE_SIZE };
            hr = factory->GetImage(size, SIIGBF_BIGGERSIZEOK, bmp);
        }
        return hr;
    }

    bool SetObject(IDataObject* dataObject) throw()
    {
        if (dataObject == NULL || DataObject::IsMyDataObject(dataObject)) {
            return false;
        }
        this->dataObject.Release();
        if (this->bitmap != NULL) {
            DeleteBitmap(this->bitmap);
            this->bitmap = NULL;
        }
        if (this->bitmapPM != NULL) {
            DeleteBitmap(this->bitmapPM);
            this->bitmapPM = NULL;
        }
        if (this->name != NULL) {
            ::SHFree(this->name);
            this->name = NULL;
        }
        if (FAILED(DataObject::Clone(dataObject, &this->dataObject))) {
            return false;
        }
        HRESULT hr = GetBitmap(this->dataObject, &this->bitmap);
        if (SUCCEEDED(hr)) {
            SIZE size;
            this->bitmapPM = Premultiply(this->bitmap, &size);
            this->width = size.cx;
            this->height = size.cy;
            ComPtr<IShellItem> item;
            if (SUCCEEDED(::SHGetItemFromObject(this->dataObject, IID_PPV_ARGS(&item)))) {
                item->GetDisplayName(SIGDN_NORMALDISPLAY, &this->name);
            }
        }
        HWND hwnd = NULL;
        if (SUCCEEDED(this->GetWindow(&hwnd))) {
            ::InvalidateRect(hwnd, NULL, FALSE);
        }
        return true;
    }

    static HBITMAP Premultiply(__in HBITMAP bmp, __out SIZE* size) throw()
    {
        BITMAP bm = {};
        ::GetObject(bmp, sizeof(bm), &bm);
        const int width = bm.bmWidth;
        const int height = bm.bmHeight;
        size->cx = width;
        size->cy = height;
        if (width <= 0 || height <= 0) {
            return NULL;
        }
        HDC hdc = ::CreateCompatibleDC(NULL);
        DWORD* image;
        BITMAPINFO bi = { { sizeof(bi.bmiHeader), width, -height, 1, 32 } };
        HBITMAP bmpPM = ::CreateDIBSection(hdc, &bi, DIB_RGB_COLORS, reinterpret_cast<void**>(&image), NULL, 0);
        if (bmpPM != NULL) {
            if (::GetDIBits(hdc, bmp, 0, height, image, &bi, DIB_RGB_COLORS) == height) {
                size_t size = width * height;
                DWORD* ptr = image;
                while (size--) {
                    BYTE b = *ptr & 0xff;
                    BYTE g = (*ptr >> 8) & 0xff;
                    BYTE r = (*ptr >> 16) & 0xff;
                    BYTE a = (*ptr >> 24) & 0xff;
                    b = ((static_cast<UINT>(b) * a) >> 8) & 0xff;
                    g = ((static_cast<UINT>(g) * a) >> 8) & 0xff;
                    r = ((static_cast<UINT>(r) * a) >> 8) & 0xff;
                    *ptr = (a << 24) | (r << 16) | (g << 8) | b;
                    ptr++;
                }
            }
        }
        ::DeleteDC(hdc);
        return bmpPM;
    }

public:
    MainWindow() throw()
        : toolbar()
        , closeImage()
        , bitmap()
        , bitmapPM()
        , width()
        , height()
        , font()
        , textHeight()
        , name()
    {
    }

    ~MainWindow() throw()
    {
        if (this->closeImage != NULL) {
            ::ImageList_Destroy(this->closeImage);
        }
        if (this->bitmap != NULL) {
            DeleteBitmap(this->bitmap);
        }
        if (this->bitmapPM != NULL) {
            DeleteBitmap(this->bitmapPM);
        }
        if (this->font != NULL) {
            DeleteFont(this->font);
        }
        if (this->name != NULL) {
            ::SHFree(this->name);
        }
    }

    HWND Create() throw()
    {
        ::InitCommonControls();
        const ATOM atom = __super::Register(MAKEINTATOM(411));
        if (atom == 0) {
            return NULL;
        }
        return __super::Create(atom, NULL, WS_POPUP | WS_SYSMENU | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_TOOLWINDOW);
    }

    void AdjustWindowPos(HWND hwnd, int width, int height) throw()
    {
        RECT rect = { 0, 0, width, height + this->textHeight };
        MyAdjustWindowRect(hwnd, &rect);
        CenterWindowPos(hwnd, HWND_TOPMOST, rect.right - rect.left, rect.bottom - rect.top, SWP_SHOWWINDOW);
        ::UpdateWindow(hwnd);
    }

    void SetItem(IShellItem* item) throw()
    {
        if (item == NULL) {
            return;
        }
        ComPtr<IDataObject> dataObject;
        if (SUCCEEDED(item->BindToHandler(NULL, BHID_DataObject, IID_PPV_ARGS(&dataObject)))) {
            this->SetObject(dataObject);
        }
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

class CBufferedPaintInit
{
public:
    HRESULT m_hr;

    CBufferedPaintInit() throw()
        : m_hr(::BufferedPaintInit())
    {
    }

    ~CBufferedPaintInit() throw()
    {
        if (SUCCEEDED(m_hr)) {
            ::BufferedPaintUnInit();
        }
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

int Main(int argc, __in_ecount(argc + 1) LPWSTR* argv) throw()
{
#ifdef USETRACE
    CAllocConsole con;
#endif
    COleInitialize ole;
    CBufferedPaintInit bp;
    MainWindow wnd;
    HWND hwnd = wnd.Create();
    if (hwnd == NULL) {
        return EXIT_FAILURE;
    }
    wnd.AdjustWindowPos(hwnd, 256, 256);
    if (argc > 1) {
        ComPtr<IShellItem> item;
        if (FAILED(::SHCreateItemFromParsingName(argv[1], NULL, IID_PPV_ARGS(&item)))) {
            WCHAR fullpath[MAX_PATH];
            if (::GetFullPathNameW(argv[1], ARRAYSIZE(fullpath), fullpath, NULL)) {
                ::SHCreateItemFromParsingName(fullpath, NULL, IID_PPV_ARGS(&item));
            }
        }
        wnd.SetItem(item);
    }
    MSG msg;
    msg.wParam = EXIT_FAILURE;
    while (::GetMessage(&msg, NULL, 0, 0)) {
        ::DispatchMessage(&msg);
        ::TranslateMessage(&msg);
    }
    return static_cast<int>(msg.wParam);
}

int WINAPI wWinMain(HINSTANCE, HINSTANCE, LPWSTR, int) throw()
{
    int argc = 0;
    LPWSTR* argv = ::CommandLineToArgvW(::GetCommandLineW(), &argc);
    if (argv == NULL) {
        return 255;
    }
    return Main(argc, argv);
}

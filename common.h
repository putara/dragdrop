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

#pragma warning(disable: 4127)
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

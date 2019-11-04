#include "common.h"
#include "ComHelper.h"
#include "Array.h"
#include "Utils.h"
#include "BaseWindow.h"
#include "DataObject.h"
#include "DropTarget.h"
#include "Resize.h"
#include "Magnet.h"
#include "MainWindow.h"
#include "Init.h"


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
    wnd.AdjustWindowPos(hwnd, 256, 256);
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

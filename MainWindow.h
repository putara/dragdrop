
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
    static const int BTN_MARGIN = 4;

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
        HANDLE_MSG(hwnd, WM_GETMINMAXINFO, OnGetMinMaxInfo);
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
                ::SetWindowPos(this->toolbar, NULL, rect.right - btnRect.right - BTN_MARGIN, BTN_MARGIN, btnRect.right, btnRect.bottom, 0);
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

    void OnGetMinMaxInfo(HWND, LPMINMAXINFO minmax) throw()
    {
        minmax->ptMinTrackSize.x = BTN_SIZE + BTN_MARGIN * 2;
        minmax->ptMinTrackSize.y = BTN_SIZE + BTN_MARGIN * 2;
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
        bool ret = this->SetObject(dataObject);
        HWND hwnd = NULL;
        if (SUCCEEDED(this->GetWindow(&hwnd))) {
            ::InvalidateRect(hwnd, NULL, TRUE);
        }
        if (ret) {
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
        ComPtr<IShellItem> item;
        if (SUCCEEDED(::SHGetItemFromDataObject(dataObject, DOGIF_TRAVERSE_LINK, IID_PPV_ARGS(&item)))) {
            item->GetDisplayName(SIGDN_NORMALDISPLAY, &this->name);
        }
        HRESULT hr = GetBitmap(this->dataObject, &this->bitmap);
        if (SUCCEEDED(hr)) {
            SIZE size;
            this->bitmapPM = Premultiply(this->bitmap, &size);
            this->width = size.cx;
            this->height = size.cy;
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


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

    static bool IsBlacklisted(const FORMATETC& fetc) throw()
    {
        if ((fetc.dwAspect & DVASPECT_CONTENT) == 0) {
            return true;
        }
        WCHAR name[MAX_PATH];
        if (::GetClipboardFormatNameW(fetc.cfFormat, name, ARRAYSIZE(name)) == 0) {
            return false;
        }
        static const WCHAR c_excludes[] =
            L"UsingDefaultDragImage\0"
            L"DragContext\0"
            L"DragWindow\0"
            L"DisableDragText\0"
            L"ComputedDragImage\0"
            L"InShellDragLoop\0"
            L"PersistedDataObject\0";
        LPCWSTR p = c_excludes;
        while (*p != 0) {
            if (::wcscmp(name, p) == 0) {
                return true;
            }
            p += ::wcslen(p) + 1;
        }
        return false;
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
                    if (IsBlacklisted(item.fetc)) {
                        continue;
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

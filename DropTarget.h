
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

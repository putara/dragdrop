#ifndef COMHELPER_H
#define COMHELPER_H

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

#endif

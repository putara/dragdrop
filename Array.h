#ifndef ARRAY_H
#define ARRAY_H

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

#endif

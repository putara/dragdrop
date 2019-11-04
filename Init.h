
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

// DataObject.h


STDAPI CDataObject_CreateInstance(REFIID riid, void **ppv);


//////////////////////////////////////////////////////
// CDataObject

class CDataObject : public IDataObject
{
private:
	long _cRef;

    CComPtr<IDataObject> _pDataObjShell;

public:
	CDataObject();
private:
	~CDataObject();

	HRESULT _EnsureShellDataObject();

public:

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
	IFACEMETHODIMP_(ULONG) AddRef();
	IFACEMETHODIMP_(ULONG) Release();

	// IDataObject
	IFACEMETHODIMP GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium);
	IFACEMETHODIMP GetDataHere(FORMATETC* /* pformatetc */, STGMEDIUM* /* pmedium */);
	IFACEMETHODIMP QueryGetData(FORMATETC *pformatetc);
	IFACEMETHODIMP GetCanonicalFormatEtc(FORMATETC *pformatetcIn, FORMATETC *pFormatetcOut);
    IFACEMETHODIMP SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease);
    IFACEMETHODIMP EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc);
	IFACEMETHODIMP DAdvise(FORMATETC* /* pformatetc */, DWORD /* advf */, IAdviseSink* /* pAdvSnk */, DWORD* /* pdwConnection */);
	IFACEMETHODIMP DUnadvise(DWORD /* dwConnection */);
	IFACEMETHODIMP EnumDAdvise(IEnumSTATDATA** /* ppenumAdvise */);
};
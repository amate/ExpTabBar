// DataObject.cpp

#include "stdafx.h"
#include <strsafe.h>
#include <ObjIdl.h>

#include "DataObject.h"


STDAPI CDataObject_CreateInstance(REFIID riid, void **ppv)
{
    *ppv = NULL;
    CDataObject *p = new (std::nothrow) CDataObject();
    HRESULT hr = p ? S_OK : E_OUTOFMEMORY;
    if (SUCCEEDED(hr))
    {
        hr = p->QueryInterface(riid, ppv);
        p->Release();
    }
    return hr;
}

//////////////////////////////////////////////////////
// CDataObject

CDataObject::CDataObject() : _cRef(1), _pDataObjShell(NULL)
{
}

CDataObject::~CDataObject()
{
}

HRESULT CDataObject::_EnsureShellDataObject()
{
	return _pDataObjShell ? S_OK : SHCreateDataObject(NULL, 0, NULL, NULL, IID_PPV_ARGS(&_pDataObjShell));
}


// IUnknown
IFACEMETHODIMP CDataObject::QueryInterface(REFIID riid, void **ppv)    {
    static const QITAB qit[] = {
        QITABENT(CDataObject, IDataObject),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) CDataObject::AddRef()
{
	return InterlockedIncrement(&_cRef);
}

IFACEMETHODIMP_(ULONG) CDataObject::Release()
{
	long cRef = InterlockedDecrement(&_cRef);
	if (0 == cRef) {
	    delete this;
	}
	return cRef;
}

// IDataObject
IFACEMETHODIMP CDataObject::GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
{
	ZeroMemory(pmedium, sizeof(*pmedium));

	HRESULT hr = DATA_E_FORMATETC;
	if (pformatetcIn->cfFormat == CF_UNICODETEXT)
	{
	    if (pformatetcIn->tymed & TYMED_HGLOBAL)
	    {
			WCHAR const c_szText[] = L"“ú–{Œê‚Å‚¨‚‹";
	        UINT cb = sizeof(c_szText[0]) * (lstrlen(c_szText) + 1);
	        HGLOBAL h = GlobalAlloc(GPTR, cb);
	        hr = h ? S_OK : E_OUTOFMEMORY;
	        if (SUCCEEDED(hr))
	        {
	            StringCbCopy((PWSTR)h, cb, c_szText);
	            pmedium->hGlobal = h;
	            pmedium->tymed = TYMED_HGLOBAL;
	        }
	    }
	}
	else if (SUCCEEDED(_EnsureShellDataObject()))
	{
	    hr = _pDataObjShell->GetData(pformatetcIn, pmedium);
	}
	return hr;
}

IFACEMETHODIMP CDataObject::GetDataHere(FORMATETC* /* pformatetc */, STGMEDIUM* /* pmedium */)
{
	return E_NOTIMPL;
}

IFACEMETHODIMP CDataObject::QueryGetData(FORMATETC *pformatetc)
{
    HRESULT hr = S_FALSE;
    if (pformatetc->cfFormat == CF_UNICODETEXT)
    {
        hr = S_OK;
    }
    else if (SUCCEEDED(_EnsureShellDataObject()))
    {
        hr = _pDataObjShell->QueryGetData(pformatetc);
    }
    return hr;
}


IFACEMETHODIMP CDataObject::GetCanonicalFormatEtc(FORMATETC *pformatetcIn, FORMATETC *pFormatetcOut)
{
    *pFormatetcOut = *pformatetcIn;
    pFormatetcOut->ptd = NULL;
    return DATA_S_SAMEFORMATETC;
}

IFACEMETHODIMP CDataObject::SetData(FORMATETC *pformatetc, STGMEDIUM *pmedium, BOOL fRelease)
{
    HRESULT hr = _EnsureShellDataObject();
    if (SUCCEEDED(hr))
    {
        hr = _pDataObjShell->SetData(pformatetc, pmedium, fRelease);
    }
    return hr;
}

IFACEMETHODIMP CDataObject::EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc)
{
    *ppenumFormatEtc = NULL;
    HRESULT hr = E_NOTIMPL;
    if (dwDirection == DATADIR_GET)
    {
        FORMATETC rgfmtetc[] =
        {
            // the order here defines the accuarcy of rendering
            { CF_UNICODETEXT,          NULL, 0,  0, TYMED_HGLOBAL },
        };
        hr = SHCreateStdEnumFmtEtc(ARRAYSIZE(rgfmtetc), rgfmtetc, ppenumFormatEtc);
    }
    return hr;
}



IFACEMETHODIMP CDataObject::DAdvise(FORMATETC* /* pformatetc */, DWORD /* advf */, IAdviseSink* /* pAdvSnk */, DWORD* /* pdwConnection */)
{
	return E_NOTIMPL;
}

IFACEMETHODIMP CDataObject::DUnadvise(DWORD /* dwConnection */)
{
	return E_NOTIMPL;
}

IFACEMETHODIMP CDataObject::EnumDAdvise(IEnumSTATDATA** /* ppenumAdvise */)
{
	return E_NOTIMPL;
}



























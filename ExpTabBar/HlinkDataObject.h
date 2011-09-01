// @file	HlinkDataObject.h
// @brief	ハイパーリンク・データ・オブジェクト

#pragma once

#include "DonutFunc.h"


namespace MTL
{

// for debug
#ifdef _DEBUG
	const bool _Mtl_HlinkDataObject_traceOn = false;
	#define HLDTRACE	if (_Mtl_HlinkDataObject_traceOn) ATLTRACE
#else
	#define HLDTRACE
#endif



/////////////////////////
// MtlCom.hから

LPFORMATETC _MtlFillFormatEtc(LPFORMATETC lpFormatEtc, CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtcFill);

bool MtlIsDataAvailable(IDataObject *pDataObject, CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc = NULL);


// cf. "ATL Internals"
__inline HRESULT WINAPI _This(void *pv, REFIID iid, void **ppvObject, DWORD_PTR)
{
	ATLASSERT(iid == IID_NULL);
	*ppvObject = pv;
	return S_OK;
}

/////////////////////////



//////////////////////////////////////////////////////////////////////////////////
// CHlinkDataObjectBase

class CHlinkDataObjectBase 
{
public:
	static CString *	s_pstrTmpDir;
private:
	static bool 		s_bStaticInit;

public:
	// コンストラクタ
	CHlinkDataObjectBase()
	{
		// init static variables
		if (!s_bStaticInit) {
//			::EnterCriticalSection(&_Module.m_csStaticDataInit);

			if (!s_bStaticInit) {
				s_pstrTmpDir  = new CString;

				// calc tmp directory name
				*s_pstrTmpDir = Misc::GetExeDirectory() + _T("ShortcutTmp\\");
				HLDTRACE(_T(" s_pstrTmpDir(%s)\n"), *s_pstrTmpDir);

				// done
				s_bStaticInit = true;
			}

//			::LeaveCriticalSection(&_Module.m_csStaticDataInit);
		}
	}

	static void _CreateTmpDirectory()
	{
		HLDTRACE( _T("CHlinkDataObjectBase::_CreateTmpDirectory\n") );
		ATLASSERT(s_pstrTmpDir != NULL);
		// create directory
		SECURITY_ATTRIBUTES sa;
		sa.nLength				= sizeof (SECURITY_ATTRIBUTES);
		sa.lpSecurityDescriptor = NULL;
		sa.bInheritHandle		= FALSE;
		::CreateDirectory(*s_pstrTmpDir, &sa);
	}


	// call when your application finished
	static void Term()
	{
		_RemoveTmpDirectory();

		if (s_pstrTmpDir)
			delete s_pstrTmpDir;

		s_pstrTmpDir = NULL;
	}


	static void _CleanCurTmpFiles()
	{
		if (s_pstrTmpDir == NULL)	// if no DataObject had been ever created,
			return;

		MtlForEachFile( *s_pstrTmpDir, _Function_RemoveFile() );
	}


private:
	struct _Function_RemoveFile {
		void operator ()(const CString &strFile)
		{
			if ( MtlIsExt( strFile, _T(".url") ) )
				::DeleteFile(strFile);
		}
	};


	static void _RemoveTmpDirectory()
	{
		if (s_pstrTmpDir == NULL)	// if no DataObject had been ever created,
			return;

		MtlForEachFile( *s_pstrTmpDir, _Function_RemoveFile() );
		::RemoveDirectory(*s_pstrTmpDir);
	}

};


// 初期化
__declspec(selectany) bool 		CHlinkDataObjectBase::s_bStaticInit	= false;
__declspec(selectany) CString*	CHlinkDataObjectBase::s_pstrTmpDir 	= NULL;




//////////////////////////////////////////////////////////////////////////////
// CHlinkDataObject

class ATL_NO_VTABLE CHlinkDataObject
	: public CHlinkDataObjectBase
	, public CComCoClass<CHlinkDataObject, &CLSID_NULL>
	, public CComObjectRootEx<CComSingleThreadModel>
	, public IDataObjectImpl<CHlinkDataObject>
{
public:
	DECLARE_NO_REGISTRY()
	DECLARE_NOT_AGGREGATABLE(CHlinkDataObject)

	// Data members
	CComPtr<IDataAdviseHolder>					m_spDataAdviseHolder;
	CSimpleArray< std::pair<CString, CString> >	m_arrNameAndUrl;

private:
	bool										m_bInitialized;
	CSimpleArray<CString>						m_arrFileNames;

public:
	BEGIN_COM_MAP(CHlinkDataObject)
		COM_INTERFACE_ENTRY(IDataObject)
		COM_INTERFACE_ENTRY_FUNC(IID_NULL, 0, _This)
	END_COM_MAP()

	// コンストラクタ
	CHlinkDataObject() : m_bInitialized(false)
	{
		HLDTRACE( _T("CHlinkDataObject::CHlinkDataObject\n") );
	}

	// Overrides
	HRESULT	IDataObject_GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium)
	{
		HLDTRACE( _T("CHlinkDataObject::IDataObject_GetData\n") );

		if (pformatetcIn == NULL || pmedium == NULL)
			return E_POINTER;

		::memset( pmedium, 0, sizeof (STGMEDIUM) );

		if ( (pformatetcIn->tymed & TYMED_HGLOBAL) == 0 )
			return DATA_E_FORMATETC;

		pmedium->tymed			= TYMED_HGLOBAL;
		pmedium->pUnkForRelease = NULL;

		HGLOBAL hGlobal = NULL;

		if ( pformatetcIn->cfFormat == CF_UNICODETEXT 
		  || pformatetcIn->cfFormat == ::RegisterClipboardFormat(CFSTR_SHELLURL) ) {
			hGlobal = _CreateText();
		} else if (pformatetcIn->cfFormat == CF_HDROP) {
			// next, create shortcuts
			_InitFileNamesArrayForHDrop();
			// third, create dropfile
			hGlobal = (HGLOBAL) MtlCreateDropFile(m_arrFileNames);
			ATLASSERT(hGlobal != NULL);
		}

		if (hGlobal != NULL) {
			pmedium->hGlobal = hGlobal;
			return S_OK;
		} else
			return S_FALSE;
	}

private:
	///////////////////////
	// IDataObject
	STDMETHOD	(QueryGetData) (FORMATETC * pformatetc)
	{
		if (pformatetc == NULL)
			return E_POINTER;

		if ( pformatetc->cfFormat == CF_HDROP 
		  || pformatetc->cfFormat == CF_UNICODETEXT
		  || pformatetc->cfFormat == ::RegisterClipboardFormat(CFSTR_SHELLURL) )
			return NOERROR;
		else
			return DATA_E_FORMATETC;
	}


	STDMETHOD	(EnumFormatEtc) (DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc)
	{
		HLDTRACE( _T("CHlinkDataObject::EnumFormatEtc\n") );

		if (ppenumFormatEtc == NULL)
			return E_POINTER;

		if (dwDirection == DATADIR_SET)
			return S_FALSE;

		typedef CComEnum< IEnumFORMATETC, &IID_IEnumFORMATETC, FORMATETC, _Copy<FORMATETC> >  CEnumFormatEtc;

		CComObject<CEnumFormatEtc>* pEnumFormatEtc;
		HRESULT	hr = CComObject<CEnumFormatEtc>::CreateInstance(&pEnumFormatEtc);

		if ( FAILED(hr) ) {
			HLDTRACE( _T(" CComObject<CEnumFormatEtc>::CreateInstance failed\n") );
			return S_FALSE;
		}

		pEnumFormatEtc->AddRef();	// preparing for the failure of the following QueryInterface

		{
			FORMATETC formatetcs[] = {
				{ CF_HDROP, 								 NULL,		DVASPECT_CONTENT,	-1,   TYMED_HGLOBAL },
				{ CF_UNICODETEXT,							 NULL,		DVASPECT_CONTENT,	-1,   TYMED_HGLOBAL },
				{ ::RegisterClipboardFormat(CFSTR_SHELLURL), NULL,		DVASPECT_CONTENT,	-1,   0 			},
			};
			hr = pEnumFormatEtc->Init(formatetcs, formatetcs + _countof(formatetcs), NULL, AtlFlagCopy);
			hr = pEnumFormatEtc->QueryInterface(IID_IEnumFORMATETC, (void **) ppenumFormatEtc);
		}

		pEnumFormatEtc->Release();

		return hr;
	}

	// Implementation
	void _InitFileNamesArrayForHDrop()
	{
		// on demmand
		if (m_bInitialized)
			return;

		ATLASSERT(m_arrFileNames.GetSize() == 0);
		_CreateTmpDirectory();		// on demmand

		// first, clear prev tmp files now
		_CleanCurTmpFiles();

		HLDTRACE( _T("CHlinkDataObject::_InitFileNamesArrayForHDrop\n") );

		CSimpleArray<CString> arrTmpFiles;

		for (int i = 0; i < m_arrNameAndUrl.GetSize(); ++i) {
			CString strName = m_arrNameAndUrl[i].first;
			MTL::MtlValidateFileName(strName);	// ファイル名に使えない文字を置換する
			CString strUrl	= m_arrNameAndUrl[i].second;
			HLDTRACE(_T(" (%s, %s)\n"), strName, strUrl);

			if ( !strName.IsEmpty() ) { 		// not a local file
				strName = _CompactFileName( *s_pstrTmpDir, strName, _T(".url") );
				strName = _UniqueFileName(arrTmpFiles, strName);

				if ( _CreateInternetShortcutFile(strName, strUrl) )
					arrTmpFiles.Add(strName);
				else
					ATLASSERT(FALSE);

				m_arrFileNames.Add(strName);
			} else {							// local file
				m_arrFileNames.Add(strUrl);
			}
		}

		m_bInitialized = true;
	}


	// ファイルパスの長さがMAX_PATHに収まるようにする
	static CString _CompactFileName(const CString &strDir, const CString &strFile, const CString &strExt)
	{
		// it's enough about 5
		int nRest = MAX_PATH - strDir.GetLength() - strExt.GetLength() - 5; 

		ATLASSERT( nRest > 0 && _T("Your Application path is too deep") );
		return strDir + strFile.Left(nRest) + strExt;
	}


	HGLOBAL _CreateText()
	{
		if (m_arrNameAndUrl.GetSize() == 0)
			return NULL;

		CString		strText  = m_arrNameAndUrl[0].second;
		DWORD		size     = ::lstrlen(strText) + 1;
		HGLOBAL		hMem	 = ::GlobalAlloc( GHND, size * sizeof(TCHAR) );

		if (hMem == NULL)
			return NULL;

		LPTSTR	lpszDest = (LPTSTR) ::GlobalLock(hMem);
		::lstrcpyn(lpszDest, strText, size);
		::GlobalUnlock(hMem);

		return hMem;
	}


	bool _CreateInternetShortcutFile(const CString &strFileName, const CString &strUrl)
	{
		ATLASSERT(strFileName.GetLength() <= MAX_PATH);
		HLDTRACE(_T(" _CreateInternetShortcutFile(%s)\n"), strFileName);

		return MtlCreateInternetShortcutFile(strFileName, strUrl);
	}


	CString _UniqueFileName(CSimpleArray<CString> &arrFileNames, const CString &strFileName)
	{
		CString strNewName = strFileName;
		CString strTmp;
		int 	i		   = 0;

		while (arrFileNames.Find(strNewName) != -1) {
			strTmp	   = _RemoveExtFromPath(strFileName);
			strTmp	  += _T('[');
			strTmp.AppendFormat(_T("%s%d"), strTmp, i++);
			strTmp	  += _T(']');
			strTmp	  += _GetExtFromPath(strFileName);
			strNewName = strTmp;
		}

		return strNewName;
	}

	// パスから拡張子を取り除きます
	CString _RemoveExtFromPath(const CString &strPath)
	{
		CString strResult = strPath;
		int 	nIndex	  = strPath.ReverseFind( _T('.') );

		if (nIndex != -1)
			strResult = strResult.Left(nIndex);

		return strResult;
	}

	// パスから拡張子を取りだします
	CString _GetExtFromPath(const CString &strPath)
	{
		return strPath.Right(4);
	}

};




inline bool _MtlIsHlinkDataObject(IDataObject *pDataObject)
{
	return	MtlIsDataAvailable(pDataObject, CF_HDROP)
		 || MtlIsDataAvailable(pDataObject, CF_UNICODETEXT)
		 //|| MtlIsDataAvailable(pDataObject, ::RegisterClipboardFormat(CFSTR_SHELLURL))
		 || MtlIsDataAvailable(pDataObject, ::RegisterClipboardFormat(CFSTR_SHELLIDLIST));
}



}	// namespace MTL


using namespace MTL;




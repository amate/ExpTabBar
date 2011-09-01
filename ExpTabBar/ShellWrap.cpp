/**
*	@file	ShellWrap.cpp
*	@brief	シェルに関数便利な関数
*/

#include "stdafx.h"
#include "ShellWrap.h"
#include <boost/tokenizer.hpp>


namespace ShellWrap
{

/// アイテムＩＤリストからアイコンを作る
HICON	CreateIconFromIDList(PCIDLIST_ABSOLUTE pidl)
{
	HRESULT	hr;
	CComPtr<IShellItem>	pShellItem;
	hr = ::SHCreateItemFromIDList(pidl, IID_IShellItem, (LPVOID*)&pShellItem);
	if (hr == S_OK) {
		CComPtr<IExtractIcon>	pExtractIcon;
		hr = pShellItem->BindToHandler(NULL, BHID_SFUIObject, IID_IExtractIcon, (LPVOID*)&pExtractIcon);
		if (hr == S_OK) {
			WCHAR	szPath[MAX_PATH];
			int		iIndex;
			UINT	uFlags;
			hr = pExtractIcon->GetIconLocation(GIL_FORSHELL, szPath, MAX_PATH, &iIndex, &uFlags);
			if (hr == S_OK) {
				HICON	hIcon;
				hr = pExtractIcon->Extract(szPath, iIndex, NULL, &hIcon, MAKELONG(32, 16));
				if (hr == S_OK) {
					ATLASSERT(hIcon);
					return hIcon;
				}
			}
		}
	}
	//ATLASSERT(FALSE);
	return NULL;
}


/// フルパスからアイテムＩＤリストを作成する
PIDLIST_ABSOLUTE CreateIDListFromFullPath(LPCTSTR strFullPath)
{
	LPITEMIDLIST pidl;
	if (::SHILCreateFromPath(strFullPath, &pidl, NULL) == S_OK) {
		return pidl;
	}
	ATLASSERT(FALSE);
	return NULL;
}

/// アイテムＩＤリストからフルパスを返す
CString	GetFullPathFromIDList(PCIDLIST_ABSOLUTE pidl)
{
	PWSTR	strFullPath;
	CString FullPath;
	if (::SHGetNameFromIDList(pidl, SIGDN_DESKTOPABSOLUTEPARSING, &strFullPath) == S_OK) {
		 FullPath = strFullPath;
		::CoTaskMemFree(strFullPath);
	} else {
		ATLASSERT(FALSE);
	}
	return FullPath;
}


/// アイテムＩＤリストの表示名を返す
CString GetNameFromIDList(PCIDLIST_ABSOLUTE pidl)
{
	PWSTR	strName;
	CString Name;
	if (::SHGetNameFromIDList(pidl, SIGDN_NORMALDISPLAY, &strName) == S_OK) {
		 Name = strName;
		::CoTaskMemFree(strName);
	} else {
		ATLASSERT(FALSE);
	}
	return Name;
}

/// 現在エクスプローラーで表示中のフォルダのアイテムＩＤリストを作成する
PIDLIST_ABSOLUTE GetCurIDList(IShellBrowser* pShellBrowser)
{
	try {
		PIDLIST_ABSOLUTE	pidl;
		HRESULT	hr;
		CComPtr<IShellView>	pShellView;
		CComQIPtr<IFolderView>	pFolderView;
		CComQIPtr<IPersistFolder2>	pPersistFolder2;
		hr = pShellBrowser->QueryActiveShellView(&pShellView);
		if (hr == S_OK) {
			pPersistFolder2 = pFolderView = pShellView;
			hr = pFolderView->GetFolder(IID_IPersistFolder2, (void**)&pPersistFolder2);
			if (hr == S_OK) {
				hr = pPersistFolder2->GetCurFolder(&pidl);
				if (hr == S_OK) {
					return pidl;
				}
			}
		}
//		throw AtlThrow(hr);
	}
//	catch (AtlException &) {
//		MessageBox(_T("AtlException"));
//	}
	catch (...) {
		ATLASSERT(FALSE);
	}
	return NULL;
}


/// lnkのリンク先を返します
PIDLIST_ABSOLUTE GetResolveIDList(PIDLIST_ABSOLUTE pidl)
{
	CString strLinkPath = ShellWrap::GetFullPathFromIDList(pidl);

	CComPtr<IShellLink>	spShellLink;
	HRESULT hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void **)&spShellLink);
	if (FAILED(hr)) 
		return nullptr;

	CComQIPtr<IPersistFile>	spPersistFile = spShellLink;
	hr = spPersistFile->Load(strLinkPath, STGM_READ);
	if (FAILED(hr)) 
		return nullptr;

	hr = spShellLink->Resolve(NULL, SLR_NOSEARCH);
	LPITEMIDLIST	pidlLink;
	hr = spShellLink->GetIDList(&pidlLink);
	return pidlLink;
}




/// アイテムＩＤリストで示されるフォルダが存在するかどうか
bool	IsExistFolderFromIDList(PCIDLIST_ABSOLUTE pidl)
{
	CString strFolderPath = GetFullPathFromIDList(pidl);
	DWORD	dwAttributes = ::GetFileAttributes(strFolderPath);
	if (   dwAttributes != -1 && dwAttributes & FILE_ATTRIBUTE_DIRECTORY
		|| strFolderPath.GetLength() >= 3 && (strFolderPath.Mid(1) == _T(":\\") || strFolderPath.Left(3) == _T("::{"))) 
	{
		return true;
	}
	return false;
#if 0
	CString strFolderPath = GetFullPathFromIDList(pidl);
	if (::PathIsDirectory(strFolderPath)
		|| strFolderPath.Mid(1, 2) != _T(":\\")) {
		return true;
	}
	return false;
#endif
}

LPBYTE	GetByteArrayFromBinary(int nSize, std::wstring strBinary)
{
	BYTE *pabID = new BYTE[nSize - sizeof(USHORT)];

	typedef boost::char_separator<wchar_t> Separator;
	typedef boost::tokenizer<Separator, std::wstring::const_iterator, std::wstring> Tokenizer;
	Separator sep(_T(","));
	Tokenizer tokens(strBinary, sep);

	int i = 0;
	for ( Tokenizer::iterator it = tokens.begin(); it != tokens.end(); ++it, ++i ) {
		pabID[i] = _wtoi((LPWSTR)it->c_str());
	}

	return pabID;
}


}	// namespace ShellWrap

using namespace ShellWrap;



















































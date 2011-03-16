// ShellWrap.h
#pragma once

#include <ShObjIdl.h>
#include <ShlObj.h>
#include <boost/tokenizer.hpp>

namespace ShellWrap
{

// アイテムＩＤリストからアイコンを作る
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


PIDLIST_ABSOLUTE CreateIDListFromFullPath(LPCTSTR strFullPath)
{
	LPITEMIDLIST pidl;
	if (::SHILCreateFromPath(strFullPath, &pidl, NULL) == S_OK) {
		return pidl;
	}
	ATLASSERT(FALSE);
	return NULL;
}

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

PIDLIST_ABSOLUTE GetCurIDList(CComPtr<IShellBrowser> pShellBrowser)
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

// アイテムＩＤリストで示されるフォルダが存在するかどうか
bool	IsExistFolderFromIDList(PCIDLIST_ABSOLUTE pidl)
{
	CString strFolderPath = GetFullPathFromIDList(pidl);
	if (::PathIsDirectory(strFolderPath)
		|| strFolderPath.Mid(1, 2) != _T(":\\")) {
		return true;
	}
	return false;

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








































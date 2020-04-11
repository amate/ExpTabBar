/**
*	@file	ShellWrap.cpp
*	@brief	�V�F���Ɋ֐��֗��Ȋ֐�
*/

#include "stdafx.h"
#include "ShellWrap.h"
#include <boost/tokenizer.hpp>
#include "OleDragDropTabCtrl.h"

namespace ShellWrap
{

/// �A�C�e���h�c���X�g����A�C�R�������
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


/// �t���p�X����A�C�e��ID���X�g���쐬����
/// �쐬�����A�C�e��ID���X�g�͂����Ɖ�����邱�ƁI
PIDLIST_ABSOLUTE CreateIDListFromFullPath(LPCTSTR strFullPath)
{
	LPITEMIDLIST pidl;
	if (::SHILCreateFromPath(strFullPath, &pidl, NULL) == S_OK) {
		return pidl;
	}
	return NULL;
}

/// �A�C�e���h�c���X�g����t���p�X��Ԃ�
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


/// �A�C�e���h�c���X�g�̕\������Ԃ�
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

/// ���݃G�N�X�v���[���[�ŕ\�����̃t�H���_�̃A�C�e���h�c���X�g���쐬����
PIDLIST_ABSOLUTE GetCurIDList(IShellBrowser* pShellBrowser)
{
	PIDLIST_ABSOLUTE	pidl;
	HRESULT	hr;
	CComPtr<IShellView>	pShellView;
	CComQIPtr<IFolderView>	pFolderView;
	CComQIPtr<IPersistFolder2>	pPersistFolder2;
	hr = pShellBrowser->QueryActiveShellView(&pShellView);
	pFolderView = pShellView;
	if (hr == S_OK && pFolderView) {
		hr = pFolderView->GetFolder(IID_IPersistFolder2, (void**)&pPersistFolder2);
		if (hr == S_OK && pPersistFolder2) {
			hr = pPersistFolder2->GetCurFolder(&pidl);
			if (hr == S_OK) {
				return pidl;
			}
		}
	}

	return NULL;
}

/// ���ݕ\�����̃r���[�� nIndex �ɂ���A�C�e����pidl�𓾂�
PIDLIST_ABSOLUTE GetIDListByIndex(IShellBrowser* pShellBrowser, int nIndex)
{
	ATLASSERT( pShellBrowser );

	CComPtr<IShellView>	spShellView;
	HRESULT hr = pShellBrowser->QueryActiveShellView(&spShellView);
	CComQIPtr<IFolderView>	spFolderView = spShellView;
	if (spFolderView == nullptr)
		return nullptr;

	LPITEMIDLIST pidlChild = nullptr;
	spFolderView->Item(nIndex, &pidlChild);
	if (pidlChild == nullptr)
		return nullptr;

	LPITEMIDLIST pidlFolder = ShellWrap::GetCurIDList(pShellBrowser);
	LPITEMIDLIST pidlSelectedItem = ::ILCombine(pidlFolder, pidlChild);
	::CoTaskMemFree(pidlFolder);

	return pidlSelectedItem;
}

/// path ��InfoTipText���擾���܂�
CString GetInfoTipText(LPCTSTR path)
{
	LPITEMIDLIST pidl = ShellWrap::CreateIDListFromFullPath(path);
	if (pidl) {
		CComPtr<IShellFolder>	spShellFolder;
		LPCITEMIDLIST	pidlChild = nullptr;
		if (SUCCEEDED(::SHBindToParent(pidl, IID_IShellFolder, (void**)&spShellFolder, &pidlChild))) {
			LPCITEMIDLIST pidlChildArray[] = { pidlChild };
			UINT rgfReserved = 0;
			CComPtr<IQueryInfo>	spQueryInfo;
			if (SUCCEEDED(spShellFolder->GetUIObjectOf(NULL, 1, pidlChildArray, IID_IQueryInfo, &rgfReserved, (void**)&spQueryInfo))) {
				LPWSTR strText;
				if (SUCCEEDED(spQueryInfo->GetInfoTip(QITIPF_DEFAULT, &strText))) {
					CString str = strText;
					::CoTaskMemFree((LPVOID)strText);
					::CoTaskMemFree(pidl);
					return str;
				}
			}
		}
		::CoTaskMemFree(pidl);
	}
	return CString();
}


/// lnk�̃����N���Ԃ��܂�
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

/// .lnk�t�@�C�����쐬���܂�
bool	CreateLinkFile(LPITEMIDLIST pidl, LPCTSTR saveFilePath)
{
	CComPtr<IShellLink>	spShellLink;
	HRESULT hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (void **)&spShellLink);
	if (FAILED(hr))
		return false;

	hr = spShellLink->SetIDList(pidl);
	if (FAILED(hr))
		return false;

	CComQIPtr<IPersistFile>	spPersistFile = spShellLink;
	if (spPersistFile == nullptr)
		return false;
	return SUCCEEDED(spPersistFile->Save(saveFilePath, TRUE));
}


/// �A�C�e���h�c���X�g�Ŏ������t�H���_�����݂��邩�ǂ���
bool	IsExistFolderFromIDList(PCIDLIST_ABSOLUTE pidl)
{
	CString strFolderPath = GetFullPathFromIDList(pidl);
	DWORD	dwAttributes = ::GetFileAttributes(strFolderPath);
	if (   dwAttributes != -1 && dwAttributes & FILE_ATTRIBUTE_DIRECTORY
		|| strFolderPath.GetLength() >= 3 
		&& (strFolderPath.Mid(1) == _T(":\\") || strFolderPath.Left(3) == _T("::{"))
		|| strFolderPath.Left(5) == _T("�����ꏊ:") ) 
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

/// IDataObject����ITEMIDLIST��Ԃ��܂�
std::vector<PIDLIST_ABSOLUTE>	GetIDListFromDataObject(IDataObject* pDataObject)
{
	std::vector<PIDLIST_ABSOLUTE>	vec;
	HRESULT	hr;
	FORMATETC	fmt;
	fmt.cfFormat= RegisterClipboardFormat(CFSTR_SHELLIDLIST);
	fmt.ptd		= NULL;
	fmt.dwAspect= DVASPECT_CONTENT;
	fmt.lindex	= -1;
	fmt.tymed	= TYMED_HGLOBAL;

	STGMEDIUM	medium;
	hr = pDataObject->GetData(&fmt, &medium);
	if (hr == S_OK) {
		LPIDA pida = (LPIDA)::GlobalLock(medium.hGlobal);
		LPCITEMIDLIST pParentidl = GetPIDLFolder(pida);
		vec.reserve(pida->cidl);
		for (UINT i = 0; i < pida->cidl; ++i) {
			LPCITEMIDLIST pChildIDList = GetPIDLItem(pida, i);
			LPITEMIDLIST	pidl = ::ILCombine(pParentidl, pChildIDList);
			ATLASSERT(pidl);
			vec.push_back(pidl);
		}
	}
	return vec;
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



















































// DonutTabBar.cpp

#include "stdafx.h"
#include "DonutTabBar.h"
#include <atlenc.h>
#include <UIAutomation.h>
#include <boost/thread.hpp>
#include "ShellWrap.h"
#include "resource.h"
#include "DonutFunc.h"
#include "IniFile.h"
#include "XmlFile.h"
#include "ExpTabBarOption.h"


////////////////////////////////////////////////////////////
// CNotifyWindow

// Constructor
CNotifyWindow::CNotifyWindow(CDonutTabBar* p)
	: m_pTabBar(p)
{	}

/// �O������^�u���J��
BOOL CNotifyWindow::OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct)
{
#if 0
	LPCTSTR strFullPath = (LPCTSTR)pCopyDataStruct->lpData;
#endif
	LPITEMIDLIST pidl = (LPITEMIDLIST)pCopyDataStruct->lpData;
	if (pidl == NULL)
		return FALSE;

	m_pTabBar->ExternalOpen(::ILClone(pidl));
	return TRUE;
}




////////////////////////////////////////////////////////////////////////////////
// CDonutTabBarCtrl

CDonutTabBar::CDonutTabBar()
	: m_ClickedIndex(-1)
	, m_bTabChanging(false)
	, m_bTabChanged(false)
	, m_bNavigateLockOpening(false)
	, m_bSaveAllTab(true)
	, m_hSearch(NULL)
	, m_wndNotify(this)
	, m_nInsertIndex(-1)
	, m_bDragItemIncludeFolder(false)
	, m_pExpTabBandMessageMap(nullptr)
{
	m_vecHistoryItem.reserve(20);
}

CDonutTabBar::~CDonutTabBar()
{
}

void	CDonutTabBar::Initialize(IUnknown* punk, CMessageMap* pMap)
{
	ATLASSERT(punk);
	CComQIPtr<IServiceProvider>	spServiceProvider(punk);
	ATLASSERT(spServiceProvider);


	spServiceProvider->QueryService(SID_SShellBrowser, &m_spShellBrowser);
	ATLASSERT(m_spShellBrowser);

	spServiceProvider->QueryService(IID_ITravelLogStg, &m_spTravelLogStg);
	ATLASSERT(m_spTravelLogStg);

	spServiceProvider->QueryService(SID_SSearchBoxInfo, &m_spSearchBoxInfo);
	ATLASSERT(m_spSearchBoxInfo);

	m_pExpTabBandMessageMap = pMap;
}


// �t���O�ɂ���đ}���ʒu��Ԃ�
int		CDonutTabBar::_ManageInsert(bool bLink)
{
	int  nCurSel = GetCurSel();

	if (nCurSel == -1) {
		return 0;
	} else if (bLink && CTabBarConfig::s_bAddLinkRight) {
		return nCurSel + 1;		// �����N�̓A�N�e�B�u�ȃ^�u�̉E�ɒǉ�
	} else {
		switch (CTabBarConfig::s_nAddPos) {
		case RIGHTPOS:		return GetItemCount();	// ��ԉE�Ƀ^�u��ǉ�
		case LEFTPOS:		return 0;				// ��ԍ��ɒǉ�
		case ACTIVERIGHT:	return nCurSel + 1;		// �A�N�e�B�u�ȃ^�u�̉E�ɒǉ�
		case ACTIVELEFT:	return nCurSel;			// �A�N�e�B�u�ȃ^�u�̍��ɒǉ�
		default:
			ATLASSERT(FALSE);
			return 0;
		};
	}
}



// ���ׂẴ^�u�̏���TabList.xml�ɕۑ�����
void	CDonutTabBar::SaveAllTab()
{
	if (m_bSaveAllTab == false) {
		return;
	}

	// �g���x�����O��ۑ�����
	_SaveTravelLog(GetCurSel());

	CString	TabList = Misc::GetExeDirectory() + _T("TabList.xml");

	try {
		CXmlFileWrite	xmlWrite(TabList);

		// <TabList>
		xmlWrite.WriteStartElement(L"TabList");

		for (int i = 0; i < GetItemCount(); ++i) {
			const CTabItem &Item = GetItem(i);

			xmlWrite.WriteStartElement(L"Tab");
			xmlWrite.WriteAttributeString(L"name"	 , Item.m_strItem);
			xmlWrite.WriteAttributeString(L"FullPath", Item.m_strFullPath);
			CString	strIndex;
			strIndex.Format(_T("%d"), Item.m_nSelectedIndex);
			xmlWrite.WriteAttributeString(L"index"	 , strIndex);

			CString	strNavigate;
			if (Item.m_fsState & TCISTATE_NAVIGATELOCK) {
				strNavigate = L"true";
			} else {
				strNavigate = L"false";
			}
			xmlWrite.WriteAttributeString(L"NavigateLock", strNavigate);

			xmlWrite.WriteStartElement(L"ITEMIDLIST");
			LPITEMIDLIST pidl = GetItemIDList(i);
			do {
				xmlWrite.WriteStartElement(L"ITEM");
				CString strabID = _T("\0");
				int	ByteSize = pidl->mkid.cb;
				if (ByteSize != 0) {
					ByteSize -= sizeof(USHORT);
					for (int i = 0; i < ByteSize; ++i) {
						strabID.AppendFormat(_T("%d,"), *(pidl->mkid.abID + i));
					}
				}
				xmlWrite.WriteAttributeValue(L"cb", pidl->mkid.cb);
				xmlWrite.WriteAttributeString(L"abID", strabID);
				xmlWrite.WriteFullEndElement();
				if (ByteSize == 0)
					break;
			} while (pidl = ::ILGetNext(pidl));
			xmlWrite.WriteFullEndElement();


			xmlWrite.WriteStartElement(L"TravelLog");
			xmlWrite.WriteStartElement(L"Back");
			for (int i = 0; i < Item.m_pTravelLogBack->size();++i) {
				xmlWrite.WriteStartElement(L"item");
				xmlWrite.WriteAttributeString(L"title", Item.m_pTravelLogBack->at(i).first);
				xmlWrite.WriteAttributeString(L"url"  , Item.m_pTravelLogBack->at(i).second);
				xmlWrite.WriteFullEndElement();	// </item>
			}
			xmlWrite.WriteFullEndElement();	// </Back>
			
			xmlWrite.WriteStartElement(L"Fore");
			for (int i = 0; i < Item.m_pTravelLogFore->size();++i) {
				xmlWrite.WriteStartElement(L"item");
				xmlWrite.WriteAttributeString(L"title", Item.m_pTravelLogFore->at(i).first);
				xmlWrite.WriteAttributeString(L"url"  , Item.m_pTravelLogFore->at(i).second);
				xmlWrite.WriteFullEndElement();	// </item>
			}
			xmlWrite.WriteFullEndElement();	// </Fore>
			xmlWrite.WriteFullEndElement();	// </TravelLog>

			xmlWrite.WriteFullEndElement();	// </Tab>
		}
	}
	catch (LPCTSTR strError) {
		MessageBox(strError);
	}
	catch (...) {
		MessageBox(_T("SaveAllTabList�Ɏ��s"));
	}
}


void	CDonutTabBar::RestoreAllTab()
{
	CString	TabList = Misc::GetExeDirectory() + _T("TabList.xml");
	vector<CTabItem>	vecItem;
	CTabItem			item;

	vector<std::pair<int, BYTE*> > vecIDList;

	TRAVELLOG*			pBackLog = NULL;
	TRAVELLOG*			pForeLog = NULL;


	try {
		CXmlFileRead	xmlRead(TabList);
		CString 		Element;
		XmlNodeType 	nodeType;
		while (xmlRead.Read(&nodeType) == S_OK) {
			switch (nodeType) {
			case XmlNodeType_Element: 
				{
					Element = xmlRead.GetLocalName();	// �v�f���擾
					if (Element == _T("Tab")) {
						if (xmlRead.MoveToFirstAttribute()) {
							do {
								CString strName = xmlRead.GetLocalName();
								if (strName == _T("name")) {
									item.m_strItem = xmlRead.GetValue();
								} else if (strName == _T("FullPath")) {
									item.m_strFullPath = xmlRead.GetValue();
								} else if (strName == _T("index")) {
									item.m_nSelectedIndex = _wtoi(xmlRead.GetValue());
								} else if (strName == _T("NavigateLock")) {
									if (xmlRead.GetValue() == _T("true")) {
										item.m_fsState |= TCISTATE_NAVIGATELOCK;
									} else {
										item.m_fsState &= ~TCISTATE_NAVIGATELOCK;
									}
								}
							} while (xmlRead.MoveToNextAttribute());
						}
					} else if (Element == _T("ITEMIDLIST")) {
						CString strItem;
						while (xmlRead.GetInternalElement(_T("ITEMIDLIST"), strItem)) {
							if (strItem == _T("ITEM")) {
								int nSize;
								if (xmlRead.MoveToFirstAttribute()) {
									nSize = _wtoi(xmlRead.GetValue());
									if (nSize == 0) {
										continue;
									}
									if (xmlRead.MoveToNextAttribute()) {
										std::wstring strBinary = xmlRead.GetValue();
										if (strBinary.empty())
											continue;
										LPBYTE pabID = GetByteArrayFromBinary(nSize, strBinary);
										vecIDList.push_back(std::make_pair(nSize, pabID));
									}
								}
							}
						}
					} else if (Element == _T("Back")) {
						pBackLog = new TRAVELLOG;
						CString strItem;
						/* �����ǉ����� */
						while (xmlRead.GetInternalElement(_T("Back"), strItem)) {
							if (strItem == _T("item")) {
								CString title;
								CString url;
								if (xmlRead.MoveToFirstAttribute()) {
									title = xmlRead.GetValue();
									if (xmlRead.MoveToNextAttribute()) {
										url = xmlRead.GetValue();
										pBackLog->push_back( std::make_pair(title, url) );
									}
								}
							}
						}
						item.m_pTravelLogBack = pBackLog;
						pBackLog = NULL;

					} else if (Element == _T("Fore")) {
						pForeLog = new TRAVELLOG;
						CString strItem;
						/* �����ǉ����� */
						while (xmlRead.GetInternalElement(_T("Fore"), strItem)) {
							if (strItem == _T("item")) {
								CString title;
								CString url;
								if (xmlRead.MoveToFirstAttribute()) {
									title = xmlRead.GetValue();
									if (xmlRead.MoveToNextAttribute()) {
										url = xmlRead.GetValue();
										pForeLog->push_back( std::make_pair(title, url) );
									}
								}
							}
						}
						item.m_pTravelLogFore = pForeLog;
						pForeLog = NULL;					
					}
					break;
				}
			case XmlNodeType_EndElement:
				{
					if (xmlRead.GetLocalName() == _T("Tab")) {
						// </Tab>�ɂ����̂�

						/* ITEMIDLIST���쐬 */
						int nSize = sizeof(USHORT);
						for (int i = 0; i < vecIDList.size(); ++i) {
							nSize += vecIDList[i].first;
						}

						LPITEMIDLIST pRetIDList;
						LPITEMIDLIST pNewIDList;

						CComPtr<IMalloc> spMalloc;
						::SHGetMalloc(&spMalloc);
						pNewIDList = pRetIDList = (LPITEMIDLIST)spMalloc->Alloc(nSize);
						::SecureZeroMemory(pRetIDList, nSize);

						std::size_t	nCount = vecIDList.size();
						for (int i = 0; i < nCount ; ++i) {
							pRetIDList->mkid.cb = vecIDList[i].first;
							memcpy(pRetIDList->mkid.abID, vecIDList[i].second, pRetIDList->mkid.cb - sizeof(USHORT));
							delete vecIDList[i].second;
							pRetIDList = LPITEMIDLIST((LPBYTE)pRetIDList + vecIDList[i].first);
						}

						item.m_pidl = pNewIDList;
						vecIDList.clear();	// �g���I������̂ō폜���Ă���

						vecItem.push_back(item);
					}
					break;
				}
			}
		}
		/* �^�u�ɒǉ����Ă��� */
		for (int i = 0; i < vecItem.size(); ++i) {
			AddTabItem(vecItem[i]);
		}

	}
	catch (LPCTSTR strError) {
		MessageBox(strError);
	}
	catch (...) {
		MessageBox(_T("RestoreAllTab�Ɏ��s"));
	}
}

// ������ۑ�����
void	CDonutTabBar::SaveHistory()
{
	CString	History = Misc::GetExeDirectory() + _T("History.xml");

	try {
		CXmlFileWrite	xmlWrite(History);

		// <History>
		xmlWrite.WriteStartElement(L"History");

		std::size_t	nCount = m_vecHistoryItem.size();
		for (int i = 0; i < nCount; ++i) {
			const HistoryItem &Item = m_vecHistoryItem[i];

			// <item>
			xmlWrite.WriteStartElement(L"item");
			xmlWrite.WriteAttributeString(L"Title"	 , Item.strTitle);
			xmlWrite.WriteAttributeString(L"FullPath", Item.strFullPath);

			// <ITEMIDLIST>
			xmlWrite.WriteStartElement(L"ITEMIDLIST");
			LPITEMIDLIST pidl = Item.pidl;
			do {
				xmlWrite.WriteStartElement(L"ITEM");
				CString strabID = _T("\0");
				int	ByteSize = pidl->mkid.cb;
				if (ByteSize != 0) {
					ByteSize -= sizeof(USHORT);
					for (int i = 0; i < ByteSize; ++i) {
						strabID.AppendFormat(_T("%d,"), *(pidl->mkid.abID + i));
					}
				}
				xmlWrite.WriteAttributeValue(L"cb", pidl->mkid.cb);
				xmlWrite.WriteAttributeString(L"abID", strabID);
				xmlWrite.WriteFullEndElement();
				if (ByteSize == 0)
					break;
			} while (pidl = ::ILGetNext(pidl));
			xmlWrite.WriteFullEndElement();	// </ITEMIDLIST>


			xmlWrite.WriteStartElement(L"TravelLog");
			xmlWrite.WriteStartElement(L"Back");
			for (int i = 0; i < Item.pLogBack->size();++i) {
				xmlWrite.WriteStartElement(L"item");
				xmlWrite.WriteAttributeString(L"title", Item.pLogBack->at(i).first);
				xmlWrite.WriteAttributeString(L"url"  , Item.pLogBack->at(i).second);
				xmlWrite.WriteFullEndElement();	// </item>
			}
			xmlWrite.WriteFullEndElement();	// </Back>
			
			xmlWrite.WriteStartElement(L"Fore");
			for (int i = 0; i < Item.pLogFore->size();++i) {
				xmlWrite.WriteStartElement(L"item");
				xmlWrite.WriteAttributeString(L"title", Item.pLogFore->at(i).first);
				xmlWrite.WriteAttributeString(L"url"  , Item.pLogFore->at(i).second);
				xmlWrite.WriteFullEndElement();	// </item>
			}
			xmlWrite.WriteFullEndElement();	// </Fore>
			xmlWrite.WriteFullEndElement();	// </TravelLog>

			xmlWrite.WriteFullEndElement();	// </item>
		}
	}
	catch (LPCTSTR strError) {
		MessageBox(strError);
	}
	catch (...) {
		MessageBox(_T("SaveHistory�Ɏ��s"));
	}
}

// �����𕜌�����
void	CDonutTabBar::RestoreHistory()
{
	CString	History = Misc::GetExeDirectory() + _T("History.xml");
	vector<HistoryItem>	vecItem;
	HistoryItem			item;

	vector<std::pair<int, BYTE*> > vecIDList;

	TRAVELLOG*			pBackLog = NULL;
	TRAVELLOG*			pForeLog = NULL;

	try {
		CXmlFileRead	xmlRead(History);
		CString 		Element;
		XmlNodeType 	nodeType;
		while (xmlRead.Read(&nodeType) == S_OK) {
			switch (nodeType) {
			case XmlNodeType_Element: 
				{
					Element = xmlRead.GetLocalName();	// �v�f���擾
					if (Element == _T("item")) {
						if (xmlRead.MoveToFirstAttribute()) {
							do {
								CString strName = xmlRead.GetLocalName();
								if (strName == _T("Title")) {
									item.strTitle = xmlRead.GetValue();
								} else if (strName == _T("FullPath")) {
									item.strFullPath = xmlRead.GetValue();
								}
							} while (xmlRead.MoveToNextAttribute());
						}
					} else if (Element == _T("ITEMIDLIST")) {
						CString strItem;
						while (xmlRead.GetInternalElement(_T("ITEMIDLIST"), strItem)) {
							if (strItem == _T("ITEM")) {
								int nSize;
								if (xmlRead.MoveToFirstAttribute()) {
									nSize = _wtoi(xmlRead.GetValue());
									if (nSize == 0) {
										continue;
									}
									if (xmlRead.MoveToNextAttribute()) {
										std::wstring strBinary = xmlRead.GetValue();
										LPBYTE pabID = GetByteArrayFromBinary(nSize, strBinary);
										vecIDList.push_back(std::make_pair(nSize, pabID));
									}
								}
							}
						}
					} else if (Element == _T("Back")) {
						pBackLog = new TRAVELLOG;
						CString strItem;
						/* �����ǉ����� */
						while (xmlRead.GetInternalElement(_T("Back"), strItem)) {
							if (strItem == _T("item")) {
								CString title;
								CString url;
								if (xmlRead.MoveToFirstAttribute()) {
									title = xmlRead.GetValue();
									if (xmlRead.MoveToNextAttribute()) {
										url = xmlRead.GetValue();
										pBackLog->push_back( std::make_pair(title, url) );
									}
								}
							}
						}
						item.pLogBack = pBackLog;
						pBackLog = NULL;

					} else if (Element == _T("Fore")) {
						pForeLog = new TRAVELLOG;
						CString strItem;
						/* �����ǉ����� */
						while (xmlRead.GetInternalElement(_T("Fore"), strItem)) {
							if (strItem == _T("item")) {
								CString title;
								CString url;
								if (xmlRead.MoveToFirstAttribute()) {
									title = xmlRead.GetValue();
									if (xmlRead.MoveToNextAttribute()) {
										url = xmlRead.GetValue();
										pForeLog->push_back( std::make_pair(title, url) );
									}
								}
							}
						}
						item.pLogFore = pForeLog;
						pForeLog = NULL;					
					}
					break;
				}
			case XmlNodeType_EndElement:
				{
					if (xmlRead.GetLocalName() == _T("item")) {
						// </item>�ɂ����̂�

						/* ITEMIDLIST���쐬 */
						int nSize = sizeof(USHORT);
						for (int i = 0; i < vecIDList.size(); ++i) {
							nSize += vecIDList[i].first;
						}

						LPITEMIDLIST pRetIDList;
						LPITEMIDLIST pNewIDList;

						CComPtr<IMalloc> spMalloc;
						::SHGetMalloc(&spMalloc);
						pNewIDList = pRetIDList = (LPITEMIDLIST)spMalloc->Alloc(nSize);
						::SecureZeroMemory(pRetIDList, nSize);

						std::size_t	nCount = vecIDList.size();
						for (int i = 0; i <  nCount; ++i) {
							pRetIDList->mkid.cb = vecIDList[i].first;
							if (pRetIDList->mkid.cb == 0)
								break;
							memcpy(pRetIDList->mkid.abID, vecIDList[i].second, pRetIDList->mkid.cb - sizeof(USHORT));
							delete vecIDList[i].second;
							pRetIDList = LPITEMIDLIST((LPBYTE)pRetIDList + vecIDList[i].first);
						}

						item.pidl = pNewIDList;
						vecIDList.clear();	// �g���I������̂ō폜���Ă���

						vecItem.push_back(item);
					}
					break;
				}
			}
		}
		/* �����ɒǉ����Ă��� */
		for (int i = 0; i < vecItem.size(); ++i) {
			HICON	hIcon = CreateIconFromIDList(vecItem[i].pidl);
			if (hIcon == NULL)
				hIcon = m_imgs.ExtractIcon(0);
			if (hIcon) {
				vecItem[i].hbmp = Misc::CreateBitmapFromHICON(hIcon);
				::DestroyIcon(hIcon);
			}
			m_vecHistoryItem.push_back(vecItem[i]);
		}

	}
	catch (LPCTSTR strError) {
		MessageBox(strError);
	}
	catch (...) {
		MessageBox(_T("RestoreHistory�Ɏ��s"));
	}
}


// �A�N�e�B�u�ȃ^�u���j�������Ƃ��ɌĂ΂��
// �t���O�ɂ���Ď��ɃA�N�e�B�u�ɂ��ׂ��^�u�̃C���f�b�N�X��Ԃ�
int CDonutTabBar::_ManageClose(int nActiveIndex)
{
	int	nCount	= GetItemCount();

	if (CTabBarConfig::s_bLeftActiveOnClose) {
		// ����Ƃ��A�N�e�B�u�ȃ^�u�̍����A�N�e�B�u�ɂ���
		int nNext = nActiveIndex - 1;

		if (nNext >= 0) {
			return nNext;
		} else {
			nNext = nActiveIndex + 1;

			if (nNext < nCount) {
				return nNext;
			}
		}
	} else {
		// ����Ƃ��A�N�e�B�u�ȃ^�u�̉E���A�N�e�B�u�ɂ���
		int nNext = nActiveIndex + 1;

		if (nNext < nCount) {
			return nNext;
		} else {
			nNext = nActiveIndex - 1;

			if (nNext >= 0) {
				return nNext;
			}
		}
	}
	// �S�^�u���폜���ꂽ
	return -1;
}



// �g���x�����O��ۑ�����
// bFore == TRUE�Ō��(�i�ޕ���)�ɒǉ�
bool	CDonutTabBar::_SaveTravelLog(int nIndex)
{
	if (_IsValidIndex(nIndex) == false) {
		return false;
	}
	
	auto funcSaveTravelLog = [&](BOOL bFore, TLENUMF flags) -> bool {
		HRESULT	hr = S_OK;
		CComPtr<IEnumTravelLogEntry>	pTLEnum;
		hr = m_spTravelLogStg->EnumEntries(flags, &pTLEnum);
		if (hr != S_OK || pTLEnum == NULL) {
			ATLASSERT(FALSE);
			return false;
		}

		TRAVELLOG *arrLog;
		if (bFore) {
			arrLog = GetItem(nIndex).m_pTravelLogFore;
		} else {
			arrLog = GetItem(nIndex).m_pTravelLogBack;
		}

		if (arrLog) {
			arrLog->clear();	// �O�̗v�f��S���폜����
			delete arrLog;
			arrLog = NULL;
		}
		
		arrLog = new TRAVELLOG;

		while (1) {
			CComPtr<ITravelLogEntry>	pTLEntry;
			hr = pTLEnum->Next(1, &pTLEntry, NULL);
			if (hr != S_OK) {
				pTLEnum->Reset();
				break;
			}

			LPWSTR	title = NULL;
			LPWSTR	url	  = NULL;
			pTLEntry->GetTitle(&title);
			pTLEntry->GetURL(&url);
			arrLog->push_back( make_pair( CString(title), CString(url) ) );

			::CoTaskMemFree(title);
			::CoTaskMemFree(url);
		}

		if (bFore) {
			GetItem(nIndex).m_pTravelLogFore = arrLog;
		} else {
			GetItem(nIndex).m_pTravelLogBack = arrLog;
		}
		return true;
	};

	if (!funcSaveTravelLog(TRUE	, TLEF_RELATIVE_FORE))
		return false;
	if (!funcSaveTravelLog(FALSE, TLEF_RELATIVE_BACK))
		return false;

	return true;
}

// �g���x�����O�𕜌�����
// bFore == TRUE�Ō��(�i�ޕ���)�ɒǉ�
bool	CDonutTabBar::_RestoreTravelLog(int nIndex)
{
	ATLASSERT( _IsValidIndex(nIndex) );

	HRESULT	hr;
	for (int i = 0; i < 2; ++i) {
		BOOL	bFore;
		int		nDir;
		if (i == 0) {
			bFore	= TRUE;
			nDir	= TLEF_RELATIVE_FORE;
		} else {
			bFore	= FALSE;
			nDir	= TLEF_RELATIVE_BACK;
		}

		CComPtr<IEnumTravelLogEntry> pTLEnum;
		hr = m_spTravelLogStg->EnumEntries(nDir, &pTLEnum);
		if (FAILED(hr) || pTLEnum == NULL) {
			ATLASSERT(FALSE);
			return false;
		}

		CComPtr<ITravelLogEntry> pTLEntryBase;
		TRAVELLOG *arrLog;
		if (bFore) {
			arrLog = GetItem(nIndex).m_pTravelLogFore;
		} else {
			arrLog = GetItem(nIndex).m_pTravelLogBack;
		}
		if (arrLog == NULL) {
			return true;
		}

		int	nSize = (int)arrLog->size();
		for( int i = 0; i < nSize; ++i ) {
			CComPtr<ITravelLogEntry> pTLEntry;
			hr = m_spTravelLogStg->CreateEntry(arrLog->at(i).second, arrLog->at(i).first, pTLEntryBase, !bFore, &pTLEntry);
			if (hr != S_OK || pTLEntry == NULL) {
				ATLASSERT(FALSE);
				return false;
			}

			if (pTLEntryBase) {
				pTLEntryBase.Release();
			}
			pTLEntryBase = pTLEntry;
		}
	}

	return true;
}


// ���݂̃g���x�����O��S������
bool	CDonutTabBar::_RemoveCurAllTravelLog()
{
	HRESULT				hr;

	for (int i = 0; i < 2; ++i) {
		int		nDir;
		if (i == 0) {
			nDir	= TLEF_RELATIVE_FORE;
		} else {
			nDir	= TLEF_RELATIVE_BACK;
		}

		ITravelLogEntry*	pTLEntry = NULL;

		CComPtr<IEnumTravelLogEntry> pTLEnum;
		hr = m_spTravelLogStg->EnumEntries(nDir, &pTLEnum);
		if (FAILED(hr) || pTLEnum == NULL) {
			ATLASSERT(FALSE);
			return false;
		}

		std::vector<ITravelLogEntry*> vecEntry;
		while(1) {
			hr = pTLEnum->Next(1, &pTLEntry, NULL);
			if (hr == S_OK) {
				vecEntry.push_back(pTLEntry);
				pTLEntry = NULL;
			} else {
				std::vector<ITravelLogEntry*>::reverse_iterator rit = vecEntry.rbegin();
				while (rit != vecEntry.rend()) {
					hr = m_spTravelLogStg->RemoveEntry(*rit);
					(*rit)->Release();
					++rit;
				}

				pTLEnum->Reset();
				break;
			}
		}
	}

	return true;
}

// �i�r�Q�[�g���b�N�̂��߂̊֐�
// �o�b�N���O�����������
void	CDonutTabBar::_RemoveOneBackLog()
{
	HRESULT	hr;

	CComPtr<IEnumTravelLogEntry> pTLEnum;
	hr = m_spTravelLogStg->EnumEntries(TLEF_RELATIVE_BACK, &pTLEnum);
	if (FAILED(hr) || pTLEnum == NULL) {
		ATLASSERT(FALSE);
		return;
	}

	CComPtr<ITravelLogEntry>	pTLEntry;
	hr = pTLEnum->Next(1, &pTLEntry, NULL);
	if (hr == S_OK) {
		hr = m_spTravelLogStg->RemoveEntry(pTLEntry);
		pTLEnum->Reset();
	}
}

CString	CDonutTabBar::_OneBackLog()
{
	HRESULT	hr;

	CComPtr<IEnumTravelLogEntry>	spTLEnum;
	hr = m_spTravelLogStg->EnumEntries(TLEF_RELATIVE_BACK, &spTLEnum);
	if (FAILED(hr)) {
		return CString();
	}

	CComPtr<ITravelLogEntry>	spTLEntry;
	hr = spTLEnum->Next(1, &spTLEntry, NULL);
	if (spTLEntry == NULL) {
		return CString();
	}
	LPWSTR	Title;
	spTLEntry->GetTitle(&Title);
	CString strTitle(Title);
	::CoTaskMemFree(Title);
	spTLEnum->Reset();

	return strTitle;
}

// pidl�Ƀ}�b�`����C���f�b�N�X��Ԃ�
int		CDonutTabBar::_IDListIsEqualIndex(LPITEMIDLIST pidl)
{
	for (int i = 0; i < GetItemCount(); ++i) {
		if (::ILIsEqual(pidl, GetItemIDList(i))) {
			return i;
		}
	}
	return -1;
}


/// �����ȃ^�u���폜����
void	CDonutTabBar::_threadVoidTabRemove()
{
	::Sleep(500);
	CLockRedraw Lock(m_hWnd);
	for (int i = m_arrDragItems.GetSize() -1; i >= 0 ; --i) {
		if (IsExistFolderFromIDList(GetItemIDList(m_arrDragItems[i])) == false) {
			DeleteItem(m_arrDragItems[i]);
		}
	}
}

void	CDonutTabBar::_AddHistory(int nDestroy)
{
	if (IsExistFolderFromIDList(GetItemIDList(nDestroy)) == false) {
		return ;
	}
	/* �g���x�����O�ۑ� */
	_SaveTravelLog(nDestroy);

	HistoryItem	item;

	item.pidl		= ::ILClone(GetItemIDList(nDestroy));
	item.strTitle	= GetItemText(nDestroy);
	item.strFullPath= GetItemFullPath(nDestroy);
	HICON	hIcon = CreateIconFromIDList(item.pidl);
	item.hbmp		= Misc::CreateBitmapFromHICON(hIcon);
	::DestroyIcon(hIcon);
	item.pLogBack	= GetItem(nDestroy).m_pTravelLogBack;
	GetItem(nDestroy).m_pTravelLogBack = NULL;
	item.pLogFore	= GetItem(nDestroy).m_pTravelLogFore;
	GetItem(nDestroy).m_pTravelLogFore = NULL;


	m_vecHistoryItem.insert(m_vecHistoryItem.begin(), item);

	// �d�����폜����
	for (int i = 1; i < m_vecHistoryItem.size(); ++i) {
		if (m_vecHistoryItem[i].strFullPath == item.strFullPath) {
			_DeleteHistory(i);
			return;
		}
	}

	if (m_vecHistoryItem.size() > CTabBarConfig::s_nMaxHistoryCount) {
		_DeleteHistory(CTabBarConfig::s_nMaxHistoryCount);
	}
}

void	CDonutTabBar::_DeleteHistory(int nIndex)
{
	HistoryItem &item = m_vecHistoryItem[nIndex];
	::CoTaskMemFree(item.pidl);
	delete item.pLogBack;
	delete item.pLogFore;
	::DeleteObject(item.hbmp);
	m_vecHistoryItem.erase(m_vecHistoryItem.begin() + nIndex);
}


void	CDonutTabBar::_ClearSearchText()
{
	try {
		HRESULT	hr;
		CComPtr<IUIAutomation>	spUIAutomation;
		hr = CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&spUIAutomation);
		ATLASSERT(spUIAutomation);

		CComPtr<IUIAutomationElement>	spUIElement;
		hr = spUIAutomation->ElementFromHandle(m_hSearch, &spUIElement);
		if (FAILED(hr))
			AtlThrow(hr);
#if 1
		{
			/* �����o�[��[�����{�b�N�X]��T�� */
			VARIANT varProp;
			varProp.vt = VT_BSTR;
			varProp.bstrVal = ::SysAllocString(_T("�����{�b�N�X"));

			CComPtr<IUIAutomationCondition>	spCondition;
			hr = spUIAutomation->CreatePropertyCondition(UIA_NamePropertyId, varProp, &spCondition);
			if (FAILED(hr)) 
				AtlThrow(hr);

			::SysFreeString(varProp.bstrVal);
			::VariantClear(&varProp);

			CComPtr<IUIAutomationElement>	spUIEdit;
			hr = spUIElement->FindFirst(TreeScope_Children, spCondition, &spUIEdit);
			if (spUIEdit == NULL)
				AtlThrow(hr);

			CComPtr<IUIAutomationValuePattern>	spUIValue;
			hr = spUIEdit->GetCurrentPatternAs(UIA_ValuePatternId, IID_IUIAutomationValuePattern, (void**)&spUIValue);
			if (spUIValue == NULL)
				AtlThrow(hr);

			CComBSTR	strValue;
			spUIValue->get_CurrentValue(&strValue);
			if (strValue.Length() != 0) {
				CComBSTR	strNone(_T(""));
				spUIValue->SetValue(strNone);
				
				INPUT	inputs[2] = { 0 };
				inputs[0].type	= INPUT_KEYBOARD;
				inputs[0].ki.wVk	= VK_RETURN;
				inputs[0].ki.wScan	= MapVirtualKey(VK_RETURN, 0);
				inputs[0].ki.dwFlags= 0;

				inputs[1].type	= INPUT_KEYBOARD;
				inputs[1].ki.wVk	= VK_RETURN;
				inputs[1].ki.wScan	= MapVirtualKey(VK_RETURN, 0);
				inputs[1].ki.dwFlags= KEYEVENTF_KEYUP;
				::SendInput(2, inputs, sizeof(INPUT));
			}
		}
#endif

#if 0
		{
			/* �����o�[��[�N���A]�{�^����T�� */
			VARIANT varProp;
			varProp.vt = VT_BSTR;
			varProp.bstrVal = ::SysAllocString(_T("�N���A"));

			CComPtr<IUIAutomationCondition>	spCondition;
			hr = spUIAutomation->CreatePropertyCondition(UIA_NamePropertyId, varProp, &spCondition);
			if (FAILED(hr)) 
				AtlThrow(hr);

			::VariantClear(&varProp);

			CComPtr<IUIAutomationElement>	spUIEClear;
			hr = spUIElement->FindFirst(TreeScope_Children, spCondition, &spUIEClear);
			if (spUIEClear == NULL)
				AtlThrow(hr);

			CComPtr<IUIAutomationInvokePattern>	spUIInvoke;
			hr = spUIEClear->GetCurrentPatternAs(UIA_InvokePatternId, IID_IUIAutomationInvokePattern, (void**)&spUIInvoke);
			if (FAILED(hr))
				AtlThrow(hr);

			spUIInvoke->Invoke();

			INPUT	inputs[2] = { 0 };
			inputs[0].type	= INPUT_KEYBOARD;
			inputs[0].ki.wVk	= VK_RETURN;
			inputs[0].ki.wScan	= MapVirtualKey(VK_RETURN, 0);
			inputs[0].ki.dwFlags= 0;

			inputs[1].type	= INPUT_KEYBOARD;
			inputs[1].ki.wVk	= VK_RETURN;
			inputs[1].ki.wScan	= MapVirtualKey(VK_RETURN, 0);
			inputs[1].ki.dwFlags= KEYEVENTF_KEYUP;
			::SendInput(2, inputs, sizeof(INPUT));
		}
#endif

	}
	catch (const CAtlException& e) {
		e;
	}
}


// �V�����^�u�����
// bAddLast == true�ōŌ�ɒǉ�
// bInset == true�Őݒ肳�ꂽ�}���ʒu(m_nInsertIndex)�ɒǉ�
int		CDonutTabBar::OnTabCreate(LPITEMIDLIST pidl, bool bAddLast /*= false*/, bool bInsert /*= false*/, bool bLink /*= false*/)
{
	ATLASSERT(pidl);

	int	nPos;
	if (bAddLast) {
		nPos = GetItemCount();
	} else {
		nPos = _ManageInsert(bLink);
	}
	if (bInsert && m_nInsertIndex != -1)
		nPos = m_nInsertIndex;

	CTabItem	item;

	item.m_pidl	= pidl;

	item.m_strItem = GetNameFromIDList(pidl);

	HICON hIcon = CreateIconFromIDList(pidl);
	item.m_nImgIndex = AddIcon(hIcon);

	item.m_strFullPath = GetFullPathFromIDList(pidl);

	item.m_pTravelLogBack = new TRAVELLOG;
	item.m_pTravelLogFore = new TRAVELLOG;

	int nNewIndex = InsertItem(nPos, item);
	if (bLink && CTabBarConfig::s_bLinkActive)
		SetCurSel(nNewIndex);
	return nNewIndex;
}

// nDestroy�ɂ���^�u���폜����
void	CDonutTabBar::OnTabDestroy(int nDestroyIndex)
{
	ATLASSERT( _IsValidIndex(nDestroyIndex) );

	/* �i�r�Q�[�g���b�N���ꂽ�^�u�͕��Ȃ� */
	if (GetItemState(nDestroyIndex) & TCISTATE_NAVIGATELOCK) {
		return;
	}

	/* �ŋߕ����^�u�ɒǉ����� */
	_AddHistory(nDestroyIndex);

	int nCurIndex = GetCurSel();
	if (nCurIndex == nDestroyIndex) {
		// �A�N�e�B�u�ȃr���[���j�����ꂽ
		// ���̃^�u���A�N�e�B�u�ɂ���
		int nNextIndex = _ManageClose(nCurIndex);
		if (nNextIndex == -1) {
			// �Ō�̃^�u������ꂽ
			return ;	// ���̂Ƃ�����Ȃ�
		}
		SetCurSel(nNextIndex);

		DeleteItem(nCurIndex);			// �^�u���폜����
	} else {
		// �A�N�e�B�u�łȂ��r���[���j�����ꂽ
		DeleteItem(nDestroyIndex);
	}
}



int		CDonutTabBar::AddTabItem(CTabItem& item)
{
#if 0
	if (item.m_pidl == NULL) {
		item.m_pidl = CreateIDListFromFullPath(item.m_strFullPath);
		if (item.m_pidl == NULL) {
			return -1;
		}
	}
#endif

	HICON hIcon = CreateIconFromIDList(item.m_pidl);
	if (hIcon) {
		item.m_nImgIndex = AddIcon(hIcon);
	} else {
		item.m_nImgIndex = 0;
	}
	return AddItem(item);
}

void	CDonutTabBar::NavigateLockTab(int nIndex, bool bOn)
{
	ATLASSERT( _IsValidIndex(nIndex) );

	CTabItem& item = GetItem(nIndex);
	if (bOn) {
		item.ModifyState(0, TCISTATE_NAVIGATELOCK); 
	} else {
		item.ModifyState(TCISTATE_NAVIGATELOCK, 0);
	}
	InvalidateRect(item.m_rcItem, FALSE);
}


void	CDonutTabBar::ExternalOpen(LPITEMIDLIST pidl)
{
	Misc::SetForegroundWindow(GetTopLevelWindow());
	int nIndex = _IDListIsEqualIndex(pidl);
	if (nIndex != -1) {					
		if (nIndex != GetCurSel()) {		/* ���łɊJ���Ă�^�u������΂������Ɉړ����� */
			_SaveSelectedIndex(GetCurSel());
			SetCurSel(nIndex);
		}
	} else {
		SetCurSel(OnTabCreate(pidl));
	}
}

void	CDonutTabBar::ExternalOpen(LPCTSTR strFullPath)
{
	LPITEMIDLIST pidl = CreateIDListFromFullPath(strFullPath);
	ExternalOpen(pidl);
}


// ���ۂɃ^�u��؂�ւ���
void	CDonutTabBar::OnSetCurSel(int nIndex, int nOldIndex)
{
	m_bTabChanging = true;		// RefreshTab��false

	if (nOldIndex != -1) {
//		SaveItemIndex(nOldIndex);
		
		if (m_bNavigateLockOpening) {
			//_RemoveOneBackLog();
		}
		/* �g���x�����O��ۑ� */
		_SaveTravelLog(nOldIndex);
		/* �ۑ������̂ō폜���� */
		_RemoveCurAllTravelLog();
	}

	// �r���[��؂�ւ���
	if (m_bNavigateLockOpening == false) {
		LPITEMIDLIST	pidl = GetItemIDList(nIndex);
		if (IsExistFolderFromIDList(pidl)) {
//			_ClearSearchText();

			_SaveSelectedIndex(nOldIndex);
			
			m_spShellBrowser->BrowseObject(pidl, SBSP_NOAUTOSELECT | SBSP_CREATENOHISTORY);

#if 0
			try {
				CComPtr<IShellView>	spShellView;
				HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
				if (FAILED(hr))
					AtlThrow(hr);

				HWND	hWnd;
				hr = spShellView->GetWindow(&hWnd);
				if (FAILED(hr) && hWnd == NULL)
					AtlThrow(hr);

				::SetFocus(hWnd);
			}
			catch (const CAtlException& e) {
				CString str;
				str.FormatMessage(_T("NavigateComplete2\nGetWindow�Ɏ��s : %x"), e.m_hr);
				MessageBox(str);
				ATLASSERT(FALSE);
			}
#endif
			_RestoreTravelLog(nIndex);
		} else {
			/* �t�H���_�����݂��Ȃ��̂Ń^�u����폜���� */
			GetItem(nIndex).ModifyState(TCISTATE_NAVIGATELOCK, 0);
			OnTabDestroy(nIndex);
		}
	}
}

// Drag&Drop�ɕK�v�ȃf�[�^��n��(�n���f�[�^�����낦�邽�߂̊֐�)
HRESULT		CDonutTabBar::OnGetTabCtrlDataObject(CSimpleArray<int> arrIndex, IDataObject **ppDataObject)
{
	HRESULT	hr;
	CComPtr<IShellItemArray>	spShellItemArray;
	LPITEMIDLIST*	arrpidl = new LPITEMIDLIST[arrIndex.GetSize()];
	for (int i = 0; i < arrIndex.GetSize(); ++i) {
		arrpidl[i] = GetItemIDList(arrIndex[i]);
	}

	hr = SHCreateShellItemArrayFromIDLists(arrIndex.GetSize(), (LPCITEMIDLIST*)arrpidl, &spShellItemArray);
	delete[] arrpidl;
	arrpidl = NULL;

	if (hr == S_OK) {
		hr = spShellItemArray->BindToHandler(NULL, BHID_DataObject, IID_PPV_ARGS(ppDataObject));
	}
	return hr;
}

void	CDonutTabBar::OnVoidTabRemove(const CSimpleArray<int>& arrCurDragItems)
{
	m_arrDragItems = arrCurDragItems;
	boost::thread	thrd(boost::bind(&CDonutTabBar::_threadVoidTabRemove, this));
}


DROPEFFECT CDonutTabBar::OnDragEnter(IDataObject *pDataObject, DWORD dwKeyState, CPoint point)
{
	m_bDragAccept = _MtlIsHlinkDataObject(pDataObject);

	if (m_bDragFromItself == false) {
		if (::GetAsyncKeyState(VK_LBUTTON) < 0) {	// �}�E�X�̍��{�^����������Ă��邩�ǂ���
			m_bLeftButton = true;
		} else {
			m_bLeftButton = false;
		}

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
			/* �h���b�O�A�C�e�����t�H���_���܂�ł��邩�ǂ��� */
			m_bDragItemIncludeFolder = false;
			LPIDA pida = (LPIDA)::GlobalLock(medium.hGlobal);
			LPCITEMIDLIST pParentidl = GetPIDLFolder(pida);
			for (UINT i = 0; i < pida->cidl; ++i) {
				LPCITEMIDLIST pChildIDList = GetPIDLItem(pida, i);
				LPITEMIDLIST	pidl = ::ILCombine(pParentidl, pChildIDList);
				ATLASSERT(pidl);
				if (ShellWrap::IsExistFolderFromIDList(pidl)) {
					::ILFree(pidl);
					m_bDragItemIncludeFolder = true;
					break;
				}
				::ILFree(pidl);
			}
			/* �h���b�O�A�C�e���� ���[�g�h���C�u(�V���[�g�J�b�g�������Ȃ�)���ǂ��� */
			m_bDragItemIsRoot = false;
			if (pida->cidl == 1) {
				LPCITEMIDLIST pChildIDList = GetPIDLItem(pida, 0);
				LPITEMIDLIST	pidl = ::ILCombine(pParentidl, pChildIDList);
				CString strDragItemPath = ShellWrap::GetFullPathFromIDList(pidl);
				::ILFree(pidl);
				m_bDragItemIsRoot = ::PathIsRoot(strDragItemPath) != 0	// �h���C�u���[�g
					|| strDragItemPath.Left(2) == _T("::")	// ����t�H���_
					|| pidl->mkid.cb == 0;					// �f�X�N�g�b�v
			}
			/* �h���b�O�A�C�e���̃h���C�u�i���o�[���擾 */
			m_DragItemDriveNumber = ::PathGetDriveNumber(GetFullPathFromIDList(pParentidl));

			::GlobalUnlock(medium.hGlobal);
			::ReleaseStgMedium(&medium);
		} else {
			m_bDragAccept	= false;	//\\+
			return DROPEFFECT_NONE;
		}

	}
	return __super::OnDragEnter(pDataObject, dwKeyState, point);
}

void	SetDropDescription(IDataObject* pDataObject, DROPEFFECT dropEffect, LPCTSTR strMessage, LPCTSTR strInsert)
{
	FORMATETC format = { 
		(CLIPFORMAT) ::RegisterClipboardFormat(CFSTR_DROPDESCRIPTION), 
		NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL
	};
	STGMEDIUM storageMedium = { 0 };
	storageMedium.tymed = TYMED_HGLOBAL;
	storageMedium.hGlobal = ::GlobalAlloc(GHND, sizeof(DROPDESCRIPTION));
	if (storageMedium.hGlobal) {
		DROPDESCRIPTION* pDD = (DROPDESCRIPTION*) ::GlobalLock(storageMedium.hGlobal);

		pDD->type = (DROPIMAGETYPE) dropEffect;
		if (strMessage)
			lstrcpyn(pDD->szMessage, strMessage, MAX_PATH);
		if (strInsert)
			lstrcpyn(pDD->szInsert, strInsert, MAX_PATH);

		::GlobalUnlock(storageMedium.hGlobal);
		pDataObject->SetData(&format, &storageMedium, TRUE);
	}
}

void	SetDropDescription(IDataObject* pDataObject, DROPEFFECT dropEffect, LPCTSTR strPath)
{
	LPCWSTR	strMessage = L"%1";
	switch (dropEffect) {
	case DROPIMAGE_COPY:	strMessage = L"%1 �փR�s�[";	break;
	case DROPIMAGE_MOVE:	strMessage = L"%1 �ֈړ�";	break;
	case DROPIMAGE_LINK:	strMessage = L"%1 �Ƀ����N���쐬";	break;
	}
	SetDropDescription(pDataObject, dropEffect, strMessage, strPath);
}


DROPEFFECT CDonutTabBar::OnDragOver(IDataObject *pDataObject, DWORD dwKeyState, CPoint point, DROPEFFECT dropOkEffect)
{
	// �O������h���b�v���ꂽ�ꍇ
	if (m_bDragFromItself == false) {
		if (m_bDragAccept == false) {
			return DROPEFFECT_NONE;
		}
		
		int nIndex = HitTest(point);
		if (nIndex == -1) {		// �t�H���_���^�u�O�Ƀh���b�O����Ă���
			_ClearInsertionEdge();
			if (m_bDragItemIncludeFolder) {	
				_DrawInsertionEdge(htOutside, nIndex);

				SetDropDescription(pDataObject, DROPIMAGE_LABEL, _T("�����Ƀ^�u���쐬"));
				return DROPEFFECT_LINK;
			} else {
				// �t�H���_�ȊO���h���b�v����Ă���@�������������ł����ɂȂɂ���邩��
				SetDropDescription(pDataObject, DROPIMAGE_INVALID, NULL, NULL);
				return DROPEFFECT_NONE;
			}
		}

		const CString& strTabFullPath = GetItemFullPath(nIndex); 
		// �^�[�Q�b�g�̃^�u������t�H���_�Ȃ̂Ŗ���
		if (strTabFullPath.Left(2) == _T("::")) {
			SetDropDescription(pDataObject, DROPIMAGE_INVALID, NULL, NULL);
			_ClearInsertionEdge();
			return DROPEFFECT_NONE;
		}

		_DrawInsertionEdge(htItem, nIndex);

		const CString strTitle = ShellWrap::GetNameFromIDList(GetItemIDList(nIndex));
		
		// �h���b�O���Ă�̂̓��[�g�t�H���_�Ȃ̂Ń����N�쐬�̂�
		if (m_bDragItemIsRoot) {
			SetDropDescription(pDataObject, DROPEFFECT_LINK, strTitle);
			return DROPEFFECT_LINK;
		}

		if ( (dwKeyState & (MK_CONTROL | MK_SHIFT)) == (MK_CONTROL | MK_SHIFT) ) {
			SetDropDescription(pDataObject, DROPEFFECT_LINK, strTitle);
			return DROPEFFECT_LINK;
		} else if (dwKeyState & MK_CONTROL) {			
			SetDropDescription(pDataObject, DROPEFFECT_COPY, strTitle);
			return DROPEFFECT_COPY;
		} else if (dwKeyState & MK_SHIFT) {
			SetDropDescription(pDataObject, DROPEFFECT_MOVE, strTitle);
			return DROPEFFECT_MOVE;
		}
		// �����h���C�u�������̂ňړ�
		if (m_DragItemDriveNumber == ::PathGetDriveNumber(strTabFullPath)) {
			SetDropDescription(pDataObject, DROPEFFECT_MOVE, strTitle);
			return DROPEFFECT_MOVE;	

		} else {	// �Ⴄ�h���C�u�������̂ŃR�s�[
			SetDropDescription(pDataObject, DROPEFFECT_COPY, strTitle);
			return DROPEFFECT_COPY;	
		}
	}

	return __super::OnDragOver(pDataObject, dwKeyState, point, dropOkEffect);
}

DROPEFFECT CDonutTabBar::OnDrop(IDataObject *pDataObject, DROPEFFECT dropEffect, DROPEFFECT dropEffectList, CPoint point)
{
	if (m_bDragFromItself == false) {	// �O������h���b�v���ꂽ
		HRESULT	hr;
		int nIndex = HitTest(point);
		if (nIndex == -1) {	// �^�u�O
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
				LPCITEMIDLIST pParentIDList = GetPIDLFolder(pida);
				for (UINT i = 0; i < pida->cidl; ++i) {
					LPCITEMIDLIST pChildIDList = GetPIDLItem(pida, i);
					LPITEMIDLIST	pidl = ::ILCombine(pParentIDList, pChildIDList);
					CComPtr<IShellItem>	spShellItem;
					hr = ::SHCreateItemFromIDList(pidl, IID_IShellItem, (LPVOID*)&spShellItem);
					if (hr == S_OK) {
						SFGAOF	attribute;
						spShellItem->GetAttributes(SFGAO_FOLDER | SFGAO_LINK, &attribute);
						if (attribute & SFGAO_LINK) {	// lnk�Ȃ��������
							LPITEMIDLIST	pidlLink = ShellWrap::GetResolveIDList(pidl);
							if (pidlLink) {
								OnTabCreate(pidlLink, true);
							}
							::CoTaskMemFree(pidl);

						} else if (attribute & SFGAO_FOLDER) {
							OnTabCreate(pidl, true);
							continue;
						}
					} else {
						::CoTaskMemFree(pidl);
					}
				}

				::GlobalUnlock(medium.hGlobal);
				::ReleaseStgMedium(&medium);
			} else {
				ATLASSERT(FALSE);
			}
			_ClearInsertionEdge();
			return dropEffect;

		// �ǂꂩ�̃^�u�Ƀh���b�v���ꂽ
		} else if (m_bLeftButton == false) {	
			// �E�{�^���Ńh���b�v���ꂽ�Ƃ��A���j���[��\������
			CComPtr<IShellFolder>	pShellFolder;
			LPCITEMIDLIST	ppidlLast;
			hr = ::SHBindToParent(GetItemIDList(nIndex), IID_IShellFolder, (LPVOID*)&pShellFolder, &ppidlLast);
			if (hr == S_OK) {
				CComPtr<IDropTarget>	pDropTarget;
				UINT	rgfReserved = 0;
				hr = pShellFolder->GetUIObjectOf(m_hWnd, 1, &ppidlLast, IID_IDropTarget, &rgfReserved, (void**)&pDropTarget);
				if (hr == S_OK && pDropTarget) {
					DWORD	KeyState = 0;
					if (::GetKeyState(VK_SHIFT) < 0) {
						KeyState |= MK_SHIFT;
					}
					if (::GetKeyState(VK_CONTROL) < 0) {
						KeyState |= MK_CONTROL;
					}
					POINTL	pt;
					::GetCursorPos((LPPOINT)&pt);
					hr = pDropTarget->DragEnter(pDataObject, KeyState, pt, &dropEffect);
					if (hr == S_OK) {
						hr = pDropTarget->DragOver(KeyState, pt, &dropEffect);
						if (hr == S_OK) {
							dropEffect = DROPEFFECT_COPY | DROPEFFECT_MOVE | DROPEFFECT_LINK;
							hr = pDropTarget->Drop(pDataObject, KeyState, pt, &dropEffect);
							if (hr == S_OK) {
								goto ESCAPE;
							}
						}
					}
				}
			}
		} else {	// ���{�^���Ńh���b�v���ꂽ�ꍇ
			pDataObject->AddRef();
			boost::thread	thrd(boost::bind(&CDonutTabBar::_threadPerformSHFileOperation, this, GetItemIDList(nIndex), pDataObject, dropEffect == DROPEFFECT_MOVE));
		}
	}
	ESCAPE:
	return __super::OnDrop(pDataObject, dropEffect, dropEffectList, point);
}

void	CDonutTabBar::OnDragLeave()
{
	SetDropDescription(m_spDataObject, DROPIMAGE_INVALID, NULL, NULL);
	__super::OnDragLeave();
}



// �y�[�W���؂�ւ�����̂Ń^�u���X�V����
void	CDonutTabBar::RefreshTab(LPCTSTR title)
{
	if (m_bTabChanging) {
		return;		// �^�u�؂�ւ����̒ʒm�Ȃ̂ŉ����ύX���Ȃ�
	}

	LPITEMIDLIST	pidl = GetCurIDList(m_spShellBrowser);
	if (pidl == NULL) {
		/* �^�u����ύX */
		if (GetCurSel() != -1)
			SetItemText(GetCurSel(), title);	// Windows\assembly �t�H���_�Ƃ��ŕK�v
		return;		// ���݂̃t�H���_�̃A�C�e���h�c���X�g���擾�ł��Ȃ�����
	}

	int nCurIndex = GetCurSel();
	if (nCurIndex == -1) {	// �^�u������Ȃ��Ȃ�G�N�X�v���[���[���N�������Ƃ���
		
		HWND hWndTarget = FindWindow(_T("ExpTabBar_NotifyWindow"), NULL);
		if (hWndTarget != NULL) {	/* ���ɂ��G�N�X�v���[���[���N�����Ă��� */			
			m_bSaveAllTab = false;
			CString strFullPath = GetFullPathFromIDList(pidl);
			if (   strFullPath.Left(5)  == _T("��������&")
				|| (CTabBarConfig::s_bMargeControlPanel == false 
				   && strFullPath.Left(40) == _T("::{26EE0668-A00A-44D7-9371-BEB064C98683}"))) 
			{
				_RefreshBandInfo();
				::CoTaskMemFree(pidl);
				return ;	// �������ʕ\���͕ʃE�B���h�E�ɂ���
			}

			// ���̃G�N�X�v���[���[��pidl���J���ꂽ���Ƃ�ʒm����
			COPYDATASTRUCT	cd = { 0 };
#if 0
			CString strFullPath = GetFullPathFromIDList(pidl);			
			Base64EncodeGetRequiredLength
			cd.lpData = (LPVOID)(LPCTSTR)strFullPath;
			cd.cbData = strFullPath.GetLength() * sizeof(TCHAR) + sizeof(TCHAR);
#endif
			UINT cbItemID = ::ILGetSize(pidl);
			cd.lpData	= (LPVOID)pidl;
			cd.cbData	= cbItemID;
			SendMessage(hWndTarget, WM_COPYDATA, NULL, (LPARAM)&cd);
			::CoTaskMemFree(pidl);

			::PostMessage(GetTopLevelWindow(), WM_CLOSE, 0, 0);
/*
			DWORD	dwProcessID = 0;
			GetWindowThreadProcessId(GetTopLevelWindow(), &dwProcessID);
			ATLASSERT(dwProcessID == 0);
			HANDLE h = ::OpenProcess(PROCESS_TERMINATE, FALSE, dwProcessID);
			ATLASSERT(h == NULL);
			::TerminateProcess(h, 0);*/
			return;
		} else {	// ���̃G�N�X�v���[���[�͂Ȃ�

			// �ʒm���󂯎�邽�߂̃E�B���h�E�����
			HWND hWndNotify = m_wndNotify.CreateEx(NULL);

			RestoreAllTab();	// �^�u�𕜌�����
			/* ���������^�u�Ɠ����^�u���J����邩�ǂ��� */
			int nIndex = _IDListIsEqualIndex(pidl);
			if (nIndex != -1) {
				nCurIndex = nIndex;	// �����^�u��������
			} else {
				nCurIndex = OnTabCreate(pidl, true);
			}

			SetCurSel(nCurIndex);
			return;
		}
	}

	/* ���łɊJ���Ă�^�u������΂������Ɉړ����� */
	int nIndex = _IDListIsEqualIndex(pidl);
	if (nIndex != -1 && nIndex != nCurIndex) {
		m_bNavigateLockOpening = true;
		_SaveSelectedIndex(nCurIndex);
		_RemoveOneBackLog();
		_AddHistory(nCurIndex);	// �ŋߕ����^�u�ɒǉ�����

		SetCurSel(nIndex);

		DeleteItem(nCurIndex);	// �폜����
		m_bNavigateLockOpening = false;
		return;
	}

	CString strFullPath = GetFullPathFromIDList(pidl);
	if (GetItemState(nCurIndex) & TCISTATE_NAVIGATELOCK) {
		// �i�r�Q�[�g���b�N����Ă���̂ŐV�K�^�u�ŊJ��
		if (strFullPath.Left(5) == _T("�����ꏊ:")) {
			// �������悤�Ƃ��Ă���̂Ń^�u�̍X�V�ɗ��߂�
			goto REFRESH;
		}
		if (GetItemFullPath(nCurIndex).Left(5) == _T("�����ꏊ:")) {
			// �������猳�ɖ߂����̂Ń^�u�̍X�V�ɗ��߂�
			if (GetItemText(nCurIndex) != _OneBackLog()) {
				// �������ʂ̃^�u���J�����킯����Ȃ��̂�
				goto REFRESH;
			}
		}
		m_bNavigateLockOpening = true;
		_SaveSelectedIndex(nCurIndex);
		_RemoveOneBackLog();
		int nIndex = OnTabCreate(pidl, false, false, true);	//�@�~�h���N���b�N�ŊJ���������ɂ���
		SetCurSel(nIndex);
		m_bNavigateLockOpening = false;
		return;
	}
	REFRESH:

	/* �^�u����ύX */
	SetItemText(nCurIndex, title/*GetNameFromIDList(pidl)*/);

	/* �A�C�R����ύX */
	//CreateIconFromIDList(pidl);
	ReplaceIcon(nCurIndex, CreateIconFromIDList(pidl));

	/* �A�C�e���h�c���X�g��ύX */
	SetItemIDList(nCurIndex, pidl);

	/* �t���p�X��ύX */
	SetItemFullPath(nCurIndex, strFullPath);
}


void	CDonutTabBar::NavigateComplete2(LPCTSTR strURL)
{
	if (m_bTabChanging) {
#if 0
		_ClearSearchText();
		INPUT	inputs[2] = { 0 };
		inputs[0].type	= INPUT_KEYBOARD;
		inputs[0].ki.wVk	= VK_RETURN;
		inputs[0].ki.wScan	= MapVirtualKey(VK_RETURN, 0);
		inputs[0].ki.dwFlags= 0;

		inputs[1].type	= INPUT_KEYBOARD;
		inputs[1].ki.wVk	= VK_RETURN;
		inputs[1].ki.wScan	= MapVirtualKey(VK_RETURN, 0);
		inputs[1].ki.dwFlags= KEYEVENTF_KEYUP;
		::SendInput(2, inputs, sizeof(INPUT));
		
		try {
			CComPtr<IShellView>	spShellView;
			HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
			if (FAILED(hr))
				AtlThrow(hr);

			HWND	hWnd;
			hr = spShellView->GetWindow(&hWnd);
			if (FAILED(hr) && hWnd == NULL)
				AtlThrow(hr);

			::SetFocus(hWnd);
		}
		catch (const CAtlException& e) {
			CString str;
			str.FormatMessage(_T("NavigateComplete2\nGetWindow�Ɏ��s : %x"), e.m_hr);
			MessageBox(str);
			ATLASSERT(FALSE);
		}
#endif
		m_bTabChanging	= false;
		m_bTabChanged	= true;
	}
	
}

void	CDonutTabBar::DocumentComplete()
{
#if 1
	if (m_bTabChanged) {
		m_bTabChanged = false;
		if (GetItemFullPath(GetCurSel()).Left(5) != _T("�����ꏊ:")) {
			_ClearSearchText();

			try {
				CComPtr<IShellView>	spShellView;
				HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
				spShellView->UIActivate(SVUIA_ACTIVATE_FOCUS);
				if (FAILED(hr))
					AtlThrow(hr);
#if 0
				return;

				HWND	hWnd;
				hr = spShellView->GetWindow(&hWnd);
				if (FAILED(hr) && hWnd == NULL)
					AtlThrow(hr);

				::SetFocus(hWnd);
#endif
			}
			catch (const CAtlException& e) {
				CString str;
				str.FormatMessage(_T("NavigateComplete2\nGetWindow�Ɏ��s : %x"), e.m_hr);
				MessageBox(str);
				ATLASSERT(FALSE);
			}
		}


		CListViewCtrl ListView;
		HWND hWnd;
		CComPtr<IShellView>	spShellView;
		HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
		if (SUCCEEDED(hr)) {
			hr = spShellView->GetWindow(&hWnd);
			if (SUCCEEDED(hr)) {
				ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
				if (ListView.m_hWnd != NULL) {
					int nPerPage = ListView.GetCountPerPage();
					ListView.EnsureVisible(GetItemSelectedIndex(GetCurSel()) + nPerPage - 1, FALSE);
				}
			}
		}


	}
#endif
}



// message map

int		CDonutTabBar::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	SetMsgHandled(FALSE);

	/* �ݒ��ǂݍ��� */
	CTabBarConfig::LoadConfig();

	CString	strFilePath	= Misc::GetExeDirectory() + _T("Setting.ini");
	CIniFileRead	pr(strFilePath, _T("TabCtrl"));
	SetTabStyle(pr.GetValue(_T("TabStyle"), SKN_TAB_STYLE_THEME));

	ReloadSkinData();

	RestoreHistory();

	CFavoritesOption::LoadConfig();

	/* �c�[���`�b�v���쐬 */
	m_tipHistroy.Create(m_hWnd);

	LPCTSTR kls[] =  { _T("WorkerW"), _T("ReBarWindow32"), _T("UniversalSearchBand"), _T("Search Box"), _T("SearchEditBoxWrapperClass"), NULL };
	const int ids[] = { 0, 0xA005, 0, 0, 0, -1 };
	m_hSearch = FindChildWindowIDRecursive(GetTopLevelWindow(), kls, ids);
	ATLASSERT(m_hSearch);

	m_menuPopup.LoadMenu(IDM_TAB);
	m_menuHistory = m_menuPopup.GetSubMenu(1).GetSubMenu(0);
	m_menuFavorites = m_menuPopup.GetSubMenu(1).GetSubMenu(1);

	SetTimer(AutoSaveTimerID, AutoSaveInterval);

	return 0;
}

void	CDonutTabBar::OnDestroy()
{
	SetMsgHandled(FALSE);

	if (m_bSaveAllTab) {
		_SaveSelectedIndex(GetCurSel());
		SaveAllTab();
		SaveHistory();
		std::size_t nCount = m_vecHistoryItem.size();
		for (int i = 0; i < nCount; ++i) {
			if (m_vecHistoryItem[i].pidl)
				::CoTaskMemFree(m_vecHistoryItem[i].pidl);
			m_vecHistoryItem[i].pidl = NULL;
			if (m_vecHistoryItem[i].hbmp) {
				::DeleteObject(m_vecHistoryItem[i].hbmp);
				m_vecHistoryItem[i].hbmp = NULL;
			}
		}
		if (m_wndNotify.IsWindow()) 
			m_wndNotify.DestroyWindow();

		m_bSaveAllTab = false;
	}

	CFavoritesOption::CleanFavoritesItem();

	KillTimer(AutoSaveTimerID);

	m_hWnd = NULL;
}


// �E�N���b�N
void	CDonutTabBar::OnRButtonUp(UINT nFlags, CPoint point)
{
	// Overrides
	int nIndex = HitTest(point);
	if (nIndex != -1) {
		_ExeCommand(nIndex, CTabBarConfig::s_RClickCommand);

	} else {	// �^�u�O
		m_tipHistroy.SetWindowPos(HWND_TOPMOST, -1, -1, -1, -1, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);

		CToolInfo	tinfo(TTF_SUBCLASS | TTF_TRACK, m_hWnd, 0, NULL, _T(""));
		m_tipHistroy.AddTool(tinfo);
		m_tipHistroy.Popup();

		m_tipHistroy.Activate(TRUE);
		CMenuHandle menu = m_menuPopup.GetSubMenu(1);
		ClientToScreen(&point);
		menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON,
							point.x, point.y, m_hWnd);
		m_tipHistroy.Activate(FALSE);
	}
}

void	CDonutTabBar::OnInitMenuPopup(CMenuHandle menuPopup, UINT nIndex, BOOL bSysMenu)
{
	if ((HMENU)menuPopup == (HMENU)m_menuHistory) {
		for (int i = m_menuHistory.GetMenuItemCount() -1; i >= 0 ; --i) {
			m_menuHistory.DeleteMenu(i, MF_BYPOSITION);
		}
		std::size_t	nCount = m_vecHistoryItem.size();
		for (int i = 0; i < nCount; ++i) {
			MENUITEMINFO	mii = { sizeof(MENUITEMINFO) };
			mii.fMask	= MIIM_ID | MIIM_CHECKMARKS | MIIM_STATE | MIIM_TYPE;
			mii.fType	= MFT_STRING;
			mii.fState	= MFS_UNCHECKED;
			mii.wID		= ID_RECENTCLOSED_FIRST + i;
			mii.dwTypeData	= m_vecHistoryItem[i].strTitle.GetBuffer();
			mii.cch			= m_vecHistoryItem[i].strTitle.GetLength();
			mii.hbmpUnchecked = m_vecHistoryItem[i].hbmp;
			m_menuHistory.InsertMenuItem(i, TRUE, &mii);
		}
	} else if (menuPopup.m_hMenu == m_menuFavorites.m_hMenu) {
		while (m_menuFavorites.GetMenuItemCount())
			m_menuFavorites.DeleteMenu(0, MF_BYPOSITION);
		int nCount = (int)CFavoritesOption::s_vecFavoritesItem.size();
		for (int i = 0; i < nCount; ++i) {
			auto& item = CFavoritesOption::s_vecFavoritesItem[i];
			bool bSep = item.strPath == FAVORITESSEPSTRING;
			MENUITEMINFO	mii = { sizeof(MENUITEMINFO) };
			mii.fMask	= MIIM_ID | MIIM_CHECKMARKS | MIIM_STATE | MIIM_TYPE;
			mii.fType	= bSep ? MFT_SEPARATOR : MFT_STRING;
			mii.fState	= MFS_UNCHECKED;
			mii.wID		= ID_FAVORITES_FIRST + i;
			mii.dwTypeData	= item.strTitle.GetBuffer();
			mii.cch			= item.strTitle.GetLength();
			mii.hbmpUnchecked = item.bmpIcon.m_hBitmap;
			m_menuFavorites.InsertMenuItem(i, TRUE, &mii);
		}
		if (nCount == 0) {
			MENUITEMINFO	mii = { sizeof(MENUITEMINFO) };
			mii.fMask	= MIIM_ID | MIIM_STATE | MIIM_TYPE;
			mii.fType	= MFT_STRING;
			mii.fState	= MFS_GRAYED;
			mii.dwTypeData	= _T("(�Ȃ�)");
			m_menuFavorites.InsertMenuItem(0, TRUE, &mii);
		}
	}
}

void	CDonutTabBar::OnMenuSelect(UINT nItemID, UINT nFlags, CMenuHandle menu)
{
	if ((HMENU)menu == (HMENU)m_menuHistory) {
		if (nFlags & MF_MOUSESELECT) {
			int i = nItemID - ID_RECENTCLOSED_FIRST;
			m_tipHistroy.UpdateTipText(m_vecHistoryItem[i].strFullPath.GetBuffer(), m_hWnd);

			/* �g���b�N���Ȃ� */
			CToolInfo	tinfo(0, m_hWnd);
			m_tipHistroy.TrackActivate(tinfo, TRUE);

			CPoint pt;
			GetCursorPos(&pt);
			m_tipHistroy.TrackPosition(pt.x, pt.y + 30);
			return;
		}
	} else if (menu.m_hMenu == m_menuFavorites.m_hMenu) {
		if (nFlags & MF_MOUSESELECT) {
			int i = nItemID - ID_FAVORITES_FIRST;
			m_tipHistroy.UpdateTipText(CFavoritesOption::s_vecFavoritesItem[i].strPath.GetBuffer(), m_hWnd);

			/* �g���b�N���Ȃ� */
			CToolInfo	tinfo(0, m_hWnd);
			m_tipHistroy.TrackActivate(tinfo, TRUE);

			CPoint pt;
			GetCursorPos(&pt);
			m_tipHistroy.TrackPosition(pt.x, pt.y + 30);
			return;
		}
	}

	/* �g���b�N���� */
	CToolInfo	tinfo(0, m_hWnd);
	m_tipHistroy.TrackActivate(tinfo, FALSE);
}

// �_�u���N���b�N
void	CDonutTabBar::OnLButtonDblClk(UINT nFlags, CPoint point)
{
	int nIndex = HitTest(point);
	if (nIndex != -1)
		_ExeCommand(nIndex, CTabBarConfig::s_DblClickCommand);
	else
		SetMsgHandled(FALSE);
}

// �z�C�[���N���b�N(�~�h���N���b�N)
void	CDonutTabBar::OnMButtonUp(UINT nFlags, CPoint point)
{
	SetMsgHandled(FALSE);

	int nIndex = HitTest(point);
	if (nIndex != -1)
		_ExeCommand(nIndex, CTabBarConfig::s_MClickCommand);
}

/// �����Z�[�u
void	CDonutTabBar::OnTimer(UINT_PTR nIDEvent)
{
	if (nIDEvent == AutoSaveTimerID) {
		SaveAllTab();
	}
}

/// �z�C�[���Ń^�u�؂�ւ�
BOOL	CDonutTabBar::OnMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	if (CTabBarConfig::s_bWheel) {
		int nIndex = GetCurSel();
		int nCount = GetItemCount();
		if (nCount <= 1)
			return 0;

		int nNext;
		if (zDelta > 0) {
			nNext  = nIndex - 1 < 0 ? nCount - 1 : nIndex - 1;
		} else {
			nNext  = (nIndex + 1 < nCount) ? nIndex + 1 : 0;
		}
		SetCurSel(nNext);
	}
	return 0;
}



void	CDonutTabBar::OnTabClose(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	CSimpleArray<int> arrCurMultiSel;
	GetCurMultiSelEx(arrCurMultiSel, m_ClickedIndex);

	for (int i = arrCurMultiSel.GetSize() -1; i >= 0; --i) {
		OnTabDestroy(arrCurMultiSel[i]);
	}
}

// �E���ɂ���^�u��S������
void	CDonutTabBar::OnRightAllClose(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int nIndex = m_ClickedIndex;
	if (nIndex != -1) {
		if (_IsValidIndex(nIndex +1) == false) {
			return;
		}
		for (int i = GetItemCount() -1; i > nIndex; --i) {
			OnTabDestroy(i);
		}
	}
}

// �����ɂ���^�u��S������
void	CDonutTabBar::OnLeftAllClose(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int nIndex = m_ClickedIndex;
	if (nIndex != -1) {
		if (_IsValidIndex(nIndex -1) == false) {
			return;
		}
		for (int i = nIndex -1; i >= 0; --i) {
			OnTabDestroy(i);
		}
	}
}

// �E�N���b�N���ꂽ�^�u�ȊO�����
void	CDonutTabBar::OnExceptCurTabClose(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int nIndex = m_ClickedIndex;
	if (nIndex != -1) {
		for (int i = GetItemCount() -1; i >= 0 ; --i) {
			if (i == nIndex) {
				continue;
			}
			OnTabDestroy(i);
		}
	}
}

/// ���̃t�H���_���J��
void	CDonutTabBar::OnOpenUpFolder(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int nIndex = m_ClickedIndex;
	if (nIndex != -1) {
		LPITEMIDLIST pidl = ::ILClone(m_items[nIndex].m_pidl);
		ILRemoveLastID(pidl);

		m_nInsertIndex = nIndex;
		int nNewIndex = OnTabCreate(pidl, false, true);
		SetCurSel(nNewIndex);
		m_nInsertIndex = -1;
	}
}

/// ���C�ɓ���ɒǉ�����
void	CDonutTabBar::OnAddFavorites(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int nIndex = m_ClickedIndex;
	if (nIndex != -1) {
		CFavoritesOption::AddFavorites(m_items[nIndex].m_pidl);
	}
}


/// �^�u���i�r�Q�[�g���b�N��Ԃɂ���
void	CDonutTabBar::OnNavigateLock(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	CSimpleArray<int> arrCurMultiSel;
	GetCurMultiSelEx(arrCurMultiSel, m_ClickedIndex);
	for (int i = arrCurMultiSel.GetSize() -1; i >= 0; --i) {
		CTabItem& item = GetItem(arrCurMultiSel[i]);
		if (GetItemState(arrCurMultiSel[i]) & TCISTATE_NAVIGATELOCK) {
			item.ModifyState(TCISTATE_NAVIGATELOCK, 0);
		} else {
			item.ModifyState(0 ,TCISTATE_NAVIGATELOCK);
		}
		InvalidateRect(item.m_rcItem, FALSE);
	}
}

/// �I�v�V�������J��
void	CDonutTabBar::OnOpenOption(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	CExpTabBarOption	dlg;
	dlg.Show(m_hWnd);

	_UpdateLayout();
}

/// �ŋߕ����^�u�𕜌�����
void	CDonutTabBar::OnClosedTabCreate(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int i = nID - ID_RECENTCLOSED_FIRST;
	//LPITEMIDLIST pidl = ::ILClone(m_vecHistoryItem[i].pidl);
	//int nIndex = OnTabCreate(pidl);
	CTabItem	item;
	item.m_pidl		= ::ILClone(m_vecHistoryItem[i].pidl);
	item.m_strItem		= m_vecHistoryItem[i].strTitle;
	item.m_strFullPath	= m_vecHistoryItem[i].strFullPath;
	item.m_pTravelLogBack	= new TRAVELLOG(*m_vecHistoryItem[i].pLogBack);
	item.m_pTravelLogFore	= new TRAVELLOG(*m_vecHistoryItem[i].pLogFore);
	AddTabItem(item);

	_DeleteHistory(i);
}

/// ���C�ɓ�����J��
void	CDonutTabBar::OnFavoritesOpen(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int i = nID - ID_FAVORITES_FIRST;
	LPITEMIDLIST	pidl = nullptr;
	const auto& item = CFavoritesOption::s_vecFavoritesItem[i];
	if (item.pidl)
		pidl = ::ILClone(CFavoritesOption::s_vecFavoritesItem[i].pidl);
	else
		pidl = ShellWrap::CreateIDListFromFullPath(item.strPath);
	if (pidl == nullptr) {
		MessageBox(item.strTitle + _T(" ���J���܂���ł���"), NULL, MB_ICONWARNING);
		return ;
	}
	OnTabCreate(pidl, true);
}

LRESULT CDonutTabBar::OnTooltipGetDispInfo(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
	/*
	LPNMTTDISPINFO pntdi = (LPNMTTDISPINFO)pnmh;
	pntdi->lpszText = m_strtipFullPath.GetBuffer();
	*/
	bHandled = FALSE;
	return 0;
}

/// �t�H���_�̕\���ʒu��ۑ�����
void CDonutTabBar::_SaveSelectedIndex(int nIndex)
{
	CListViewCtrl ListView;
	HWND hWnd;
	CComPtr<IShellView>	spShellView;
	HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
	if (SUCCEEDED(hr) && nIndex != -1) {
		hr = spShellView->GetWindow(&hWnd);
		if (SUCCEEDED(hr)) {
			ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
			if (ListView.m_hWnd != NULL) {
				SetItemSelectedIndex(nIndex , ListView.GetTopIndex());
			}
		}
	}
}




void	CDonutTabBar::_threadPerformSHFileOperation(LPITEMIDLIST pidlTo, IDataObject* pDataObject, bool bMove)
{
	::CoInitialize(NULL);

	// �����N���쐬����
	if (::GetKeyState(VK_SHIFT) < 0 && ::GetKeyState(VK_CONTROL) < 0) {
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
			CString strTargetFolder = ShellWrap::GetFullPathFromIDList(pidlTo);
			LPIDA pida = (LPIDA)::GlobalLock(medium.hGlobal);
			LPCITEMIDLIST pParentidl = GetPIDLFolder(pida);
			for (UINT i = 0; i < pida->cidl; ++i) {
				LPCITEMIDLIST pChildIDList = GetPIDLItem(pida, i);
				LPITEMIDLIST	pidl = ::ILCombine(pParentidl, pChildIDList);
				ATLASSERT(pidl);
				CString strLinkName = ShellWrap::GetNameFromIDList(pidl);
				strLinkName.Replace(L':', L'');
				ShellWrap::CreateLinkFile(pidl, strTargetFolder + _T("\\") + strLinkName + _T(".lnk"));
				::ILFree(pidl);
			}
		}
	} else {

		HRESULT	hr;
		CComPtr<IShellItem>	pShellItemTo;	// �����
		hr = ::SHCreateItemFromIDList(pidlTo, IID_PPV_ARGS(&pShellItemTo));
		if (hr == S_OK) {
			CComPtr<IFileOperation>	pFileOperation;
			hr = pFileOperation.CoCreateInstance(CLSID_FileOperation);
			if (hr == S_OK) {
				if (bMove){
					hr = pFileOperation->MoveItems(pDataObject, pShellItemTo);
				} else {
					hr = pFileOperation->CopyItems(pDataObject, pShellItemTo);
				}
				if (hr == S_OK) {
					hr = pFileOperation->PerformOperations();
				}
			}
		}		
	}
	pDataObject->Release();
	::CoUninitialize();
}

/// �ݒ肵���N���b�N�R�}���h�����s����
void	CDonutTabBar::_ExeCommand(int nIndex, int Command)
{
	switch (Command) {
	case TABCLOSE:
		m_ClickedIndex = nIndex;
		OnTabClose(0, 0, 0);
		break;

	case OPEN_UPFOLDER:
		m_ClickedIndex = nIndex;
		OnOpenUpFolder(0, 0, 0);
		break;

	case NAVIGATELOCK:
		m_ClickedIndex = nIndex;
		OnNavigateLock(0, 0, 0);
		break;

	case SHOWMENU: {
		ATLASSERT(m_menuPopup.IsMenu());
		CMenuHandle	menu = m_menuPopup.GetSubMenu(0);
		POINT	pt;
		::GetCursorPos(&pt);
		m_ClickedIndex = nIndex;
		menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON,
							pt.x, pt.y, m_hWnd);
		break;
	}
	default:
		break;
	};
}













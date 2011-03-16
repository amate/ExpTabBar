// DonutTabBar.cpp

#include "stdafx.h"
#include <UIAutomation.h>
#include "DonutTabBar.h"
#include "ShellWrap.h"

#include "resource.h"
#include "DonutFunc.h"
#include "IniFile.h"
#include "XmlFile.h"


////////////////////////////////////////////////////////////
// CNotifyWindow

// Constructor
CNotifyWindow::CNotifyWindow(CDonutTabBar* p)
	: m_pTabBar(p)
{
}

// 
BOOL CNotifyWindow::OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct)
{
	LPCTSTR strFullPath = (LPCTSTR)pCopyDataStruct->lpData;
	m_pTabBar->ExternalOpen(strFullPath);
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
{
	m_menuPopup.LoadMenu(IDM_TAB);
	m_menuHistory = m_menuPopup.GetSubMenu(1).GetSubMenu(0);

	m_vecHistoryItem.reserve(20);
}

CDonutTabBar::~CDonutTabBar()
{
#if 0
	CString	strFilePath	= Misc::GetExeDirectory() + _T("Setting.ini");
	CIniFileWrite	pr(strFilePath, _T("DonutTabBar"));
	pr.SetValue(m_dwExtendedStyle	, _T("ExtendedStyle"));
	pr.SetValue(m_nMaxHistory		, _T("MaxHistory"));
	pr.ChangeSectionName(_T("TabCtrl"));
#endif
}

void	CDonutTabBar::Initialize(IUnknown* punk)
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
}


// フラグによって挿入位置を返す
int		CDonutTabBar::_ManageInsert()
{
	int  nCurSel = GetCurSel();

	if (nCurSel == -1) {
		return 0;
	} else if (GetTabExtendedStyle() & MTB_EX_ADDLINKACTIVERIGHT) {
		// リンクはアクティブなタブの右に追加
		return nCurSel + 1;
	} else if (GetTabExtendedStyle() & MTB_EX_ADDLEFT) {
		// 一番左に追加
		return 0;
	} else if (GetTabExtendedStyle() & MTB_EX_ADDLEFTACTIVE) {
		// アクティブなタブの左に追加
		return nCurSel;
	} else if (GetTabExtendedStyle() & MTB_EX_ADDRIGHTACTIVE) {
		// アクティブなタブの右に追加
		return nCurSel + 1;
	} else {
		// 一番右にタブを追加
		return GetItemCount();
	}
}



// すべてのタブの情報をTabList.xmlに保存する
void	CDonutTabBar::SaveAllTab()
{
	if (m_bSaveAllTab == false) {
		return;
	}

	// トラベルログを保存する
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
		MessageBox(_T("SaveAllTabListに失敗"));
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
					Element = xmlRead.GetLocalName();	// 要素を取得
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
										LPBYTE pabID = GetByteArrayFromBinary(nSize, strBinary);
										vecIDList.push_back(std::make_pair(nSize, pabID));
									}
								}
							}
						}
					} else if (Element == _T("Back")) {
						pBackLog = new TRAVELLOG;
						CString strItem;
						/* 順次追加する */
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
						/* 順次追加する */
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
						// </Tab>にきたので

						/* ITEMIDLISTを作成 */
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
						vecIDList.clear();	// 使い終わったので削除しておく

						vecItem.push_back(item);
					}
					break;
				}
			}
		}
		/* タブに追加していく */
		for (int i = 0; i < vecItem.size(); ++i) {
			AddTabItem(vecItem[i]);
		}

	}
	catch (LPCTSTR strError) {
		MessageBox(strError);
	}
	catch (...) {
		MessageBox(_T("RestoreAllTabに失敗"));
	}
}

// 履歴を保存する
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
		MessageBox(_T("SaveHistoryに失敗"));
	}
}

// 履歴を復元する
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
					Element = xmlRead.GetLocalName();	// 要素を取得
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
						/* 順次追加する */
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
						/* 順次追加する */
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
						// </item>にきたので

						/* ITEMIDLISTを作成 */
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
						vecIDList.clear();	// 使い終わったので削除しておく

						vecItem.push_back(item);
					}
					break;
				}
			}
		}
		/* 履歴に追加していく */
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
		MessageBox(_T("RestoreHistoryに失敗"));
	}
}


// アクティブなタブが破棄されるときに呼ばれる
// フラグによって次にアクティブにすべきタブのインデックスを返す
int CDonutTabBar::_ManageClose(int nActiveIndex)
{
	int	nCount	= GetItemCount();

	if (GetTabExtendedStyle() & MTB_EX_LEFTACTIVEONCLOSE) {
		// 閉じるときアクティブなタブの左をアクティブにする
		int nNext = nActiveIndex - 1;

		if (nNext >= 0) {
			return nNext;
		} else {
			nNext = nActiveIndex + 1;

			if (nNext < nCount) {
				return nNext;
			}
		}
	} else if (GetTabExtendedStyle() & MTB_EX_RIGHTACTIVEONCLOSE) {
		// 閉じるときアクティブなタブの右をアクティブにする
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
	// 全タブが削除された
	return -1;
}



// トラベルログを保存する
// bFore == TRUEで後ろ(進む方向)に追加
bool	CDonutTabBar::_SaveTravelLog(int nIndex)
{
	if (_IsValidIndex(nIndex) == false) {
		return false;
	}

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

		CComPtr<IEnumTravelLogEntry>	pTLEnum;
		hr = m_spTravelLogStg->EnumEntries(nDir, &pTLEnum);
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
			arrLog->clear();	// 前の要素を全部削除する
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
	}
	return true;
}

// トラベルログを復元する
// bFore == TRUEで後ろ(進む方向)に追加
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


// 現在のトラベルログを全部消す
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

// ナビゲートロックのための関数
// バックログを一つだけ消す
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

// pidlにマッチするインデックスを返す
int		CDonutTabBar::_IDListIsEqualIndex(LPITEMIDLIST pidl)
{
	for (int i = 0; i < GetItemCount(); ++i) {
		if (::ILIsEqual(pidl, GetItemIDList(i))) {
			return i;
		}
	}
	return -1;
}


void	CDonutTabBar::_thread_ChangeSelectedItem(LPVOID p)
{
	((CDonutTabBar*)p)->_ChangeSelectedItem();
	_endthread();
}

void	CDonutTabBar::_ChangeSelectedItem()
{
	HRESULT	hr;
	LPITEMIDLIST	pCurFolderItemIDList= NULL;
	LPITEMIDLIST	pChildItemIDList	= NULL;
	LPITEMIDLIST	pSelectedItemIDList = NULL;
	CComPtr<IShellView>	pShellView;
	hr = m_spShellBrowser->QueryActiveShellView(&pShellView);
	if (hr == S_OK) {
		CComPtr<IFolderView2>	pFolderView2;
		hr = pShellView->QueryInterface(&pFolderView2);
		if (hr == S_OK) {
			int	nIndex;
			for (int i = 0; i < 4; ++i) {
				hr = pFolderView2->GetSelectedItem(0, &nIndex);
				ATLTRACE(_T(" nIndex : %d\n"), nIndex);
				if (nIndex != -1) {
					break;
				}
				::Sleep(100);
			}
			if (nIndex == -1) {
				return;
			}

			hr = pFolderView2->Item(nIndex, &pChildItemIDList);
			if (hr == S_OK) {
				pCurFolderItemIDList = GetCurIDList(m_spShellBrowser);
				pSelectedItemIDList = ::ILCombine(pCurFolderItemIDList, pChildItemIDList);

				CComPtr<IShellItem>	pShellItem;
				hr = ::SHCreateItemFromIDList(pSelectedItemIDList, IID_IShellItem, (void**)&pShellItem);
				if (hr == S_OK) {
					SFGAOF	Attribs = 0;
					hr = pShellItem->GetAttributes(SFGAO_FOLDER, &Attribs);
					if (hr == S_OK && Attribs == SFGAO_FOLDER) {	// フォルダの場合のみタブで開く
						/* 新しいタブで開く */
						int nNewIndex = OnTabCreate(pSelectedItemIDList);

						::CoTaskMemFree(pCurFolderItemIDList);
						::CoTaskMemFree(pChildItemIDList);
						return;
					}
				}
			}
		}
	}
	ATLASSERT(FALSE);
}

void	CDonutTabBar::_thread_VoidTabRemove(LPVOID p)
{
	((CDonutTabBar*)p)->_VoidTabRemove();
	_endthread();
}

void	CDonutTabBar::_VoidTabRemove()
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

	for (int i = 1; i < m_vecHistoryItem.size(); ++i) {
		if (m_vecHistoryItem[i].strFullPath == item.strFullPath) {
			_DeleteHistory(i);
			return;
		}
	}

	if (m_vecHistoryItem.size() > m_nMaxHistory) {
		_DeleteHistory(m_nMaxHistory);
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
			/* 検索バーの[検索ボックス]を探す */
			VARIANT varProp;
			varProp.vt = VT_BSTR;
			varProp.bstrVal = ::SysAllocString(_T("検索ボックス"));

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
			/* 検索バーの[クリア]ボタンを探す */
			VARIANT varProp;
			varProp.vt = VT_BSTR;
			varProp.bstrVal = ::SysAllocString(_T("クリア"));

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


// 新しいタブを作る
// bAddLast == trueで最後に追加
int		CDonutTabBar::OnTabCreate(LPITEMIDLIST pidl, bool bAddLast, bool bInsert)
{
	ATLASSERT(pidl);

	int	nPos;
	if (bAddLast) {
		nPos = GetItemCount();
	} else {
		nPos = _ManageInsert();
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

	return InsertItem(nPos, item);
}

// nDestroyにあるタブを削除する
void	CDonutTabBar::OnTabDestroy(int nDestroyIndex)
{
	ATLASSERT( _IsValidIndex(nDestroyIndex) );

	/* ナビゲートロックされたタブは閉じない */
	if (GetItemState(nDestroyIndex) & TCISTATE_NAVIGATELOCK) {
		return;
	}

	/* 最近閉じたタブに追加する */
	_AddHistory(nDestroyIndex);

	int nCurIndex = GetCurSel();
	if (nCurIndex == nDestroyIndex) {
		// アクティブなビューが破棄された
		// 次のタブをアクティブにする
		int nNextIndex = _ManageClose(nCurIndex);
		if (nNextIndex == -1) {
			// 最後のタブが閉じられた
			return ;	// 今のところ閉じない
		}
		SetCurSel(nNextIndex);

		DeleteItem(nCurIndex);			// タブを削除する
	} else {
		// アクティブでないビューが破棄された
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

void	CDonutTabBar::ExternalOpen(LPCTSTR strFullPath)
{
	LPITEMIDLIST pidl = CreateIDListFromFullPath(strFullPath);
	
	/* すでに開いてるタブがあればそっちに移動する */
	int nIndex = _IDListIsEqualIndex(pidl);
	if (nIndex != -1) {
		if (nIndex != GetCurSel()) {
			SetCurSel(nIndex);
		}
	} else {
		SetCurSel(OnTabCreate(pidl, true));
	}
}


// 実際にタブを切り替える
void	CDonutTabBar::OnSetCurSel(int nIndex, int nOldIndex)
{
	m_bTabChanging = true;		// RefreshTabでfalse

	if (nOldIndex != -1) {
//		SaveItemIndex(nOldIndex);
		
		if (m_bNavigateLockOpening) {
			_RemoveOneBackLog();
		}
		/* トラベルログを保存 */
		_SaveTravelLog(nOldIndex);
		/* 保存したので削除する */
		_RemoveCurAllTravelLog();
	}

	// ビューを切り替える
	if (m_bNavigateLockOpening == false) {
		LPITEMIDLIST	pidl = GetItemIDList(nIndex);
		if (IsExistFolderFromIDList(pidl)) {
//			_ClearSearchText();
			CListViewCtrl ListView;
			HWND hWnd;
			CComPtr<IShellView>	spShellView;
			HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
			if (SUCCEEDED(hr) && nOldIndex != -1) {
				hr = spShellView->GetWindow(&hWnd);
				if (SUCCEEDED(hr)) {
					ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
					if (ListView.m_hWnd != NULL) {
						SetItemSelectedIndex(nOldIndex , ListView.GetTopIndex());
					}
				}
			}


			
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
				str.FormatMessage(_T("NavigateComplete2\nGetWindowに失敗 : %x"), e.m_hr);
				MessageBox(str);
				ATLASSERT(FALSE);
			}
#endif
			_RestoreTravelLog(nIndex);
		} else {
			/* フォルダが存在しないのでタブから削除する */
			GetItem(nIndex).ModifyState(TCISTATE_NAVIGATELOCK, 0);
			OnTabDestroy(nIndex);
		}
	}
}

// Drag&Dropに必要なデータを渡す(渡すデータをそろえるための関数)
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
	_beginthread(_thread_VoidTabRemove, 0, this);
}


DROPEFFECT CDonutTabBar::OnDragEnter(IDataObject *pDataObject, DWORD dwKeyState, CPoint point)
{
	m_bDragAccept = _MtlIsHlinkDataObject(pDataObject);

	if (m_bDragFromItself == false) {
		if (::GetAsyncKeyState(VK_LBUTTON) < 0) {	// マウスの左ボタンが押されているかどうか
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
			LPIDA pida = (LPIDA)::GlobalLock(medium.hGlobal);
			LPCITEMIDLIST pParentidl = GetPIDLFolder(pida);

			CString str = GetFullPathFromIDList(pParentidl);
			m_strDraggingDrive = str.Left(3);

			::GlobalUnlock(medium.hGlobal);
			::ReleaseStgMedium(&medium);
		} else {
			ATLASSERT(FALSE);
			return DROPEFFECT_NONE;
		}

	}
	return __super::OnDragEnter(pDataObject, dwKeyState, point);
}



DROPEFFECT CDonutTabBar::OnDragOver(IDataObject *pDataObject, DWORD dwKeyState, CPoint point, DROPEFFECT dropOkEffect)
{
	// 外部からドロップされた場合
	if (m_bDragFromItself == false) {
		if (m_bDragAccept == false) {
			return DROPEFFECT_NONE;
		}

		int nIndex = HitTest(point);
		if (nIndex == -1)
			return DROPEFFECT_COPY;	// タブの上にいなかったので新規タブ
		
		if (dwKeyState & MK_CONTROL) {
			return DROPEFFECT_COPY;
		} else if (dwKeyState & MK_SHIFT) {
			return DROPEFFECT_MOVE;
		}

		if (m_strDraggingDrive == GetItemFullPath(nIndex).Left(3)) {
			return DROPEFFECT_MOVE;	// 同じドライブだったので移動
		} else {
			return DROPEFFECT_COPY;	// 違うドライブだったのでコピー
		}
	}

	return __super::OnDragOver(pDataObject, dwKeyState, point, dropOkEffect);
}

DROPEFFECT CDonutTabBar::OnDrop(IDataObject *pDataObject, DROPEFFECT dropEffect, DROPEFFECT dropEffectList, CPoint point)
{
	if (m_bDragFromItself == false) {	// 外部からドロップされた
		HRESULT	hr;
		int nIndex = HitTest(point);
		if (nIndex == -1) {	// タブ外
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
						spShellItem->GetAttributes(SFGAO_FOLDER, &attribute);
						if (attribute == SFGAO_FOLDER) {
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
			
			return dropEffect;
		}

		if (m_bLeftButton == false) {	// 右ボタンでドロップされたときのみ有効
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
		} else {	// 左クリックの場合
			HRESULT	hr;
			CComPtr<IShellItem>	pShellItemTo;	// 送り先
			hr = ::SHCreateItemFromIDList(GetItemIDList(nIndex), IID_PPV_ARGS(&pShellItemTo));
			if (hr == S_OK) {
				CComPtr<IFileOperation>	pFileOperation;
				hr = pFileOperation.CoCreateInstance(CLSID_FileOperation);
				if (hr == S_OK) {
					if (dropEffect == DROPEFFECT_MOVE){
						hr = pFileOperation->MoveItems(pDataObject, pShellItemTo);
					} else if (dropEffect == DROPEFFECT_COPY) {
						hr = pFileOperation->CopyItems(pDataObject, pShellItemTo);
					} else {
//						ATLASSERT(FALSE);	// 移動でもコピーでもなかったので帰る
						return dropEffect;
					}
					if (hr == S_OK) {
						hr = pFileOperation->PerformOperations();
						if (hr == S_OK) {
							goto ESCAPE;
						}
					}
				}
			}
		}
		ATLASSERT(FALSE);
	}
	ESCAPE:
	return __super::OnDrop(pDataObject, dropEffect, dropEffectList, point);
}




// ページが切り替わったのでタブを更新する
void	CDonutTabBar::RefreshTab(LPCTSTR title)
{
	if (m_bTabChanging) {
		return;		// タブ切り替え中の通知なので何も変更しない
	}

	LPITEMIDLIST	pidl = GetCurIDList(m_spShellBrowser);
	if (pidl == NULL) {
		return;		// 現在のフォルダのアイテムＩＤリストが取得できなかった
	}

	int nCurIndex = GetCurSel();
	if (nCurIndex == -1) {	// タブが一つもないならエクスプローラーが起動したとする
		
		HWND hWndTarget = FindWindow(_T("ExpTabBar_NotifyWindow"), NULL);
		if (hWndTarget != NULL) {	/* 他にもエクスプローラーが起動している */			
			m_bSaveAllTab = false;
			// 他のエクスプローラーにpidlが開かれたことを通知する
			COPYDATASTRUCT	cd = { 0 };
			CString strFullPath = GetFullPathFromIDList(pidl);
			cd.lpData = (LPVOID)(LPCTSTR)strFullPath;
			cd.cbData = strFullPath.GetLength() * sizeof(TCHAR) + sizeof(TCHAR);
			SendMessage(hWndTarget, WM_COPYDATA, NULL, (LPARAM)&cd);
			::CoTaskMemFree(pidl);

			SendMessage(GetTopLevelWindow(), WM_CLOSE, 0, 0);
/*
			DWORD	dwProcessID = 0;
			GetWindowThreadProcessId(GetTopLevelWindow(), &dwProcessID);
			ATLASSERT(dwProcessID == 0);
			HANDLE h = ::OpenProcess(PROCESS_TERMINATE, FALSE, dwProcessID);
			ATLASSERT(h == NULL);
			::TerminateProcess(h, 0);*/
			return;
		}

		// 通知を受け取るためのウィンドウを作る
		HWND hWndNotify = m_wndNotify.CreateEx(NULL);

		RestoreAllTab();	// タブを復元する
		/* 復元したタブと同じタブが開かれるかどうか */
		int nIndex = _IDListIsEqualIndex(pidl);
		if (nIndex != -1) {
			nCurIndex = nIndex;	// 同じタブがあった
		} else {
			nCurIndex = OnTabCreate(pidl, true);
		}

		SetCurSel(nCurIndex);
		return;
	}

	/* すでに開いてるタブがあればそっちに移動する */
	int nIndex = _IDListIsEqualIndex(pidl);
	if (nIndex != -1 && nIndex != nCurIndex) {
		m_bNavigateLockOpening = true;
		SetCurSel(nIndex);
		if ((GetItemState(nCurIndex) & TCISTATE_NAVIGATELOCK) == 0) {
			DeleteItem(nCurIndex);	// ナビゲートロックされていないので削除する
		}
		m_bNavigateLockOpening = false;
		return;
	}

	CString strFullPath = GetFullPathFromIDList(pidl);
	if (GetItemState(nCurIndex) & TCISTATE_NAVIGATELOCK) {
		// ナビゲートロックされているので新規タブで開く
		if (strFullPath.Left(5) == _T("検索場所:")) {
			// 検索しようとしているのでタブの更新に留める
			goto REFRESH;
		}
		if (GetItemFullPath(nCurIndex).Left(5) == _T("検索場所:")) {
			// 検索から元に戻ったのでタブの更新に留める
			if (GetItemText(nCurIndex) != _OneBackLog()) {
				// 検索結果のタブを開いたわけじゃないので
				goto REFRESH;
			}
		}
		m_bNavigateLockOpening = true;
		SetCurSel(OnTabCreate(pidl));
		m_bNavigateLockOpening = false;
		return;
	}
	REFRESH:

	/* タブ名を変更 */
	SetItemText(nCurIndex, title/*GetNameFromIDList(pidl)*/);

	/* アイコンを変更 */
	//CreateIconFromIDList(pidl);
	ReplaceIcon(nCurIndex, CreateIconFromIDList(pidl));

	/* アイテムＩＤリストを変更 */
	SetItemIDList(nCurIndex, pidl);

	/* フルパスを変更 */
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
			str.FormatMessage(_T("NavigateComplete2\nGetWindowに失敗 : %x"), e.m_hr);
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
		if (GetItemFullPath(GetCurSel()).Left(5) != _T("検索場所:")) {
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
				str.FormatMessage(_T("NavigateComplete2\nGetWindowに失敗 : %x"), e.m_hr);
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

	/* 設定を読み込み */
	CString	strFilePath	= Misc::GetExeDirectory() + _T("Setting.ini");
	CIniFileRead	pr(strFilePath, _T("DonutTabBar"));
	m_dwExtendedStyle	= pr.GetValue(_T("ExtendedStyle"), MTB_EX_DEFAULT_BITS);
	m_nMaxHistory		= pr.GetValue(_T("MaxHistory"), 16);

	pr.ChangeSectionName(_T("TabCtrl"));
	SetTabCtrlExtendedStyle	(pr.GetValue(_T("TabCtrlExtendedStyle")	, TAB2_EX_DEFAULT_BITS));
	SetTabStyle				(pr.GetValue(_T("TabStyle")				, SKN_TAB_STYLE_THEME));
	CSize	size;
	size.cx	= pr.GetValue(_T("TabSize.x"), 110);
	size.cy	= pr.GetValue(_T("TabSize.y"), 24);
	SetItemSize(size);
	ReloadSkinData();

	RestoreHistory();

	/* ツールチップを作成 */
	m_tipHistroy.Create(m_hWnd);

	LPCTSTR kls[] =  { _T("WorkerW"), _T("ReBarWindow32"), _T("UniversalSearchBand"), _T("Search Box"), _T("SearchEditBoxWrapperClass"), NULL };
	const int ids[] = { 0, 0xA005, 0, 0, 0, -1 };
	m_hSearch = FindChildWindowIDRecursive(GetTopLevelWindow(), kls, ids);
	ATLASSERT(m_hSearch);

	return 0;
}

void	CDonutTabBar::OnDestroy()
{
	SetMsgHandled(FALSE);

	if (m_bSaveAllTab) {
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

	m_hWnd = NULL;
}


// 右クリック
void	CDonutTabBar::OnRButtonUp(UINT nFlags, CPoint point)
{
	// Overrides
	int nIndex = HitTest(point);

	if (nIndex != -1) {
//		SaveAllTab();
//		NavigateLockTab(nIndex, true);
		
//		HWND hWndChild = GetTabPidl(nIndex);
//		ATLASSERT(hWndChild != NULL);

		if (GetTabExtendedStyle() & MTB_EX_RIGHTCLICKCLOSE) {
			OnTabDestroy(nIndex);
		} else if (m_menuPopup.m_hMenu) {
			CMenuHandle	menu = m_menuPopup.GetSubMenu(0);
			ClientToScreen(&point);
			m_ClickedIndex = nIndex;
			menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON,
								point.x, point.y, m_hWnd);
		}
	} else {
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

//	RestoreAllTab();
#if 0
		else if (GetTabExtendedStyle() & MTB_EX_RIGHTCLICKREFRESH) {
			::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_VIEW_REFRESH, 0);
		} else if ( (GetTabExtendedStyle() & MTB_EX_RIGHTCLICKCOMMAND) && m_nRClickCommand ) {
			::PostMessage(GetTopLevelParent(), WM_COMMAND, (WPARAM) m_nRClickCommand, 0);
		} else if (m_menuPopup.m_hMenu) {
			ClientToScreen(&point);
			CMenuHandle menu	 = m_menuPopup.GetSubMenu(0);

			DWORD			dwRLFlag = 0;
			CIniFileRead	pr( _GetFilePath( _T("Menu.ini") ), _T("Option") );
			pr.QueryValue( dwRLFlag, _T("REqualL") );

			if (dwRLFlag)
				dwRLFlag = TPM_RIGHTBUTTON;

			menu.TrackPopupMenu(TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON | dwRLFlag,
								point.x, point.y,
								m_hWndMenuOwner != NULL ? m_hWndMenuOwner : hWndChild);
		} else {																						// system menu (default)
			CMenuHandle menuSys = ::GetSystemMenu(hWndChild, FALSE);
			ClientToScreen(&point);
			m_wndMDIChildPopuping = hWndChild;
			_UpdateMenu(menuSys);
			menuSys.TrackPopupMenu(TPM_LEFTALIGN | TPM_TOPALIGN | TPM_LEFTBUTTON | TPM_RIGHTBUTTON, point.x, point.y, m_hWnd);			// owner is me!!
		}
	} else {
		::SendMessage(GetTopLevelParent(), WM_SHOW_TOOLBARMENU, 0, 0);
	}
#endif
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
	}
}

void	CDonutTabBar::OnMenuSelect(UINT nItemID, UINT nFlags, CMenuHandle menu)
{
	if ((HMENU)menu == (HMENU)m_menuHistory) {
		if (nFlags & MF_MOUSESELECT) {
			int i = nItemID - ID_RECENTCLOSED_FIRST;
			m_tipHistroy.UpdateTipText(m_vecHistoryItem[i].strFullPath.GetBuffer(), m_hWnd);

			/* トラックしない */
			CToolInfo	tinfo(0, m_hWnd);
			m_tipHistroy.TrackActivate(tinfo, TRUE);

			CPoint pt;
			GetCursorPos(&pt);
			m_tipHistroy.TrackPosition(pt.x, pt.y + 30);
			return;
		}
	}

	/* トラックする */
	CToolInfo	tinfo(0, m_hWnd);
	m_tipHistroy.TrackActivate(tinfo, FALSE);
}

// ダブルクリック
void	CDonutTabBar::OnLButtonDblClk(UINT nFlags, CPoint point)
{
	SetMsgHandled(FALSE);

	CSimpleArray<int> arrCurMultiSel;
	GetCurMultiSel(arrCurMultiSel, true);

	if (arrCurMultiSel.GetSize() > 1) {
		for (int i = 0; i < arrCurMultiSel.GetSize(); ++i) {
//			hWndChild = GetTabPidl(arrCurMultiSel[i]);

			if (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKCLOSE) {
				OnTabDestroy(arrCurMultiSel[i]);
			} else if (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKREFRESH) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_VIEW_REFRESH, 0);
			} else if (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKNLOCK) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_DOCHOSTUI_OPENNEWWIN, 0);
			} 
//			else if ( (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKCOMMAND) && m_nDClickCommand ) {
//				::PostMessage(GetTopLevelParent(), WM_COMMAND, (WPARAM) m_nDClickCommand, 0);
//			}
		}
	} else {
		int nIndex = HitTest(point);

		if (nIndex != -1) {
			if (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKCLOSE) {
				OnTabDestroy(nIndex);
			} else if (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKREFRESH) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_VIEW_REFRESH, 0);
			} else if (GetTabExtendedStyle() & MTB_EX_DOUBLECLICKNLOCK) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_DOCHOSTUI_OPENNEWWIN, 0);
			} 
//			else if ( (m_dwExtendedStyle & MTB_EX_DOUBLECLICKCOMMAND) && m_nDClickCommand ) {
//				::PostMessage(GetTopLevelParent(), WM_COMMAND, (WPARAM) m_nDClickCommand, 0);
//			}
		}
	}
}

// ホイールクリック(ミドルクリック)
void	CDonutTabBar::OnMButtonUp(UINT nFlags, CPoint point)
{
	SetMsgHandled(FALSE);

	CSimpleArray<int> arrCurMultiSel;
	GetCurMultiSel(arrCurMultiSel, true);

	if (arrCurMultiSel.GetSize() > 1) {
		for (int i = 0; i < arrCurMultiSel.GetSize(); ++i) {
			if (GetTabExtendedStyle() & MTB_EX_XCLICKCLOSE) {
				OnTabDestroy(arrCurMultiSel[i]);
			} else if (GetTabExtendedStyle() & MTB_EX_XCLICKREFRESH) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_VIEW_REFRESH, 0);
			} else if (GetTabExtendedStyle() & MTB_EX_XCLICKNLOCK) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_DOCHOSTUI_OPENNEWWIN, 0);
			} 
//			else if ( (GetTabExtendedStyle() & MTB_EX_XCLICKCOMMAND) && m_nXClickCommand ) {
////				::PostMessage(GetTopLevelParent(), WM_COMMAND, (WPARAM) m_nXClickCommand, 0);
//			}
		}
	} else {
		int nIndex = HitTest(point);

		if (nIndex != -1) {

//			hWndChild = GetTabPidl(nIndex);

			if (GetTabExtendedStyle() & MTB_EX_XCLICKCLOSE) {
				OnTabDestroy(nIndex);
			} else if (GetTabExtendedStyle() & MTB_EX_XCLICKREFRESH) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_VIEW_REFRESH, 0);
			} else if (GetTabExtendedStyle() & MTB_EX_XCLICKNLOCK) {
//				::PostMessage(hWndChild, WM_COMMAND, (WPARAM) ID_DOCHOSTUI_OPENNEWWIN, 0);
			} 
//			else if ( (GetTabExtendedStyle() & MTB_EX_XCLICKCOMMAND) && m_nXClickCommand ) {
////				::PostMessage(GetTopLevelParent(), WM_COMMAND, (WPARAM) m_nXClickCommand, 0);
//			}
		}
	}
}


LRESULT CDonutTabBar::OnNewTabButton(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	LPPOINT ppt = (LPPOINT)lParam;
	HRESULT	hr;
	CComPtr<IShellView>	spShellView;
	hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
	if (hr == S_OK) {
		HWND	hWnd;
		hr = spShellView->GetWindow(&hWnd);
		if (hr == S_OK) {
			ATLASSERT(::IsWindow(hWnd));

			CRect	rc;
			::GetWindowRect(hWnd, &rc);
			if (rc.PtInRect(*ppt) != 0) {
				INPUT	Inputs[2] = { 0 };
				Inputs[0].type			= INPUT_MOUSE;
				Inputs[0].mi.dwFlags	= MOUSEEVENTF_LEFTDOWN;

				Inputs[1].type			= INPUT_MOUSE;
				Inputs[1].mi.dwFlags	= MOUSEEVENTF_LEFTUP;
				::SendInput(2, Inputs, sizeof(INPUT));
			#if 0
				::ScreenToClient(m_hShellDefView, ppt);
				::SendMessage(m_hShellDefView, WM_PARENTNOTIFY, WM_LBUTTONDOWN, MAKELPARAM( (DWORD)ppt->x, (DWORD)ppt->y));
			#endif

				CComPtr<IShellView>	pShellView;
				hr = m_spShellBrowser->QueryActiveShellView(&pShellView);
				if (hr == S_OK) {
					CComPtr<IFolderView>	pFolderView;
					hr = pShellView->QueryInterface(&pFolderView);
					if (hr == S_OK) {
						hr = pFolderView->SelectItem(NULL, SVSI_DESELECTOTHERS);

						_beginthread(_thread_ChangeSelectedItem, NULL, this);
						return 0;
					}
				}
				ATLASSERT(FALSE);
			}
		}	
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

// 右側にあるタブを全部閉じる
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

// 左側にあるタブを全部閉じる
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

// 右クリックされたタブ以外を閉じる
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

/// 一つ上のフォルダを開く
void	CDonutTabBar::OnOpenUpFolder(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int nIndex = m_ClickedIndex;
	if (nIndex != -1) {
		LPITEMIDLIST pidl = ::ILClone(m_items[nIndex].m_pidl);
		ILRemoveLastID(pidl);

		m_nInsertIndex = nIndex + 1;
		OnTabCreate(pidl, false, true);
		m_nInsertIndex = -1;
	}
}

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

void	CDonutTabBar::OnClosedTabCreate(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	int i = nID - ID_RECENTCLOSED_FIRST;
	LPITEMIDLIST pidl = ::ILClone(m_vecHistoryItem[i].pidl);
	OnTabCreate(pidl);
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























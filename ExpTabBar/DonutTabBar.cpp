// DonutTabBar.cpp

#include "stdafx.h"
#include "DonutTabBar.h"
#include <atlenc.h>
#include <UIAutomation.h>
#include <thread>
#include <fstream>
#include <boost/tokenizer.hpp>
#include <nlohmann/json.hpp>

#include "ShellWrap.h"
#include "resource.h"
#include "DonutFunc.h"
#include "IniFile.h"
#include "XmlFile.h"
#include "ExpTabBarOption.h"
#include "Logger.h"
#include "CodeConvert.h"

using json = nlohmann::json;

namespace {

std::unique_ptr<char[]>	ConvertBase64fromITEMIDLIST(LPITEMIDLIST pidl)
{
	UINT totalPIDLSize = 0;
	PCUIDLIST_RELATIVE onePIDL = pidl;
	while (onePIDL) {
		UINT pidlSize = onePIDL->mkid.cb;
		if (pidlSize == 0) {
			totalPIDLSize += sizeof(USHORT);
			break;
		}
		totalPIDLSize += pidlSize;
		onePIDL = ILNext(onePIDL);
	}
	int base64Size = ATL::Base64EncodeGetRequiredLength(static_cast<int>(totalPIDLSize), ATL_BASE64_FLAG_NOCRLF);
	auto base64PIDL = std::make_unique<char[]>(base64Size + 1);
	BOOL ret = ATL::Base64Encode(reinterpret_cast<const BYTE*>(pidl), totalPIDLSize, base64PIDL.get(), &base64Size, ATL_BASE64_FLAG_NOCRLF);
	ATLASSERT(ret);
	base64PIDL[base64Size] = '\0';
	return base64PIDL;
}

LPITEMIDLIST ConvertITEMIDLISTfromBase64(const std::string& base64PIDL, CComPtr<IMalloc> spMalloc)
{
	int base64DecodeSize = ATL::Base64DecodeGetRequiredLength(static_cast<int>(base64PIDL.size()));
	LPITEMIDLIST  pNewIDList = (LPITEMIDLIST)spMalloc->Alloc(base64DecodeSize);
	::SecureZeroMemory(pNewIDList, base64DecodeSize);
	BOOL ret = ATL::Base64Decode(base64PIDL.c_str(), static_cast<int>(base64PIDL.size()), reinterpret_cast<BYTE*>(pNewIDList), &base64DecodeSize);
	ATLASSERT(ret);
	return pNewIDList;
}


}	// namespace



////////////////////////////////////////////////////////////////////////////////
// CDonutTabBarCtrl

CDonutTabBar::CDonutTabBar()
	: m_ClickedIndex(-1)
	, m_bTabChanging(false)
	, m_bTabChanged(false)
	, m_bNavigateLockOpening(false)
	, m_bSaveAllTab(true)
	, m_hSearch(NULL)
	, m_nInsertIndex(-1)
	, m_bDragItemIncludeFolder(false)
	, m_pExpTabBandMessageMap(nullptr)
{
	m_vecHistoryItem.reserve(20);

	m_mutex_currentPIDL.CreateMutex(L"kCurrentPIDLMutexName");
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

	m_pExpTabBandMessageMap = pMap;
}

void	CDonutTabBar::UnInitialize()
{
	if (m_spTravelLogStg)
		m_spTravelLogStg.Release();

	if (m_spShellBrowser)
		m_spShellBrowser.Release();

	m_pExpTabBandMessageMap = nullptr;
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
		case RIGHTPOS:	// ��ԉE�Ƀ^�u��ǉ�
#if 0
		{
			const int itemCount = GetItemCount();
			for (int i = nCurSel; i < itemCount; ++i) {
				if (GetItem(i).m_fsState & TCISTATE_LINEBREAK) {
					return i;
				}
			}
		}
#endif
		return GetItemCount();	

		case LEFTPOS:	// ��ԍ��ɒǉ�
#if 0
		{
			const int itemCount = GetItemCount();
			for (int i = nCurSel; i >= 0; --i) {
				if (GetItem(i).m_fsState & TCISTATE_LINEBREAK) {
					return i + 1;
				}
			}
		}
#endif
		return 0;

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

	json tabListArray = json::array();
	for (int i = 0; i < GetItemCount(); ++i) {
		const CTabItem& Item = GetItem(i);
		json jsonTabData;
		jsonTabData["Name"] = CodeConvert::UTF8fromUTF16((LPCWSTR)Item.m_strItem);
		if ((Item.m_fsState & TCISTATE_LINEBREAK) == 0) {
			jsonTabData["FullPath"] = CodeConvert::UTF8fromUTF16((LPCWSTR)Item.m_strFullPath);
			jsonTabData["SelectedIndex"] = Item.m_nSelectedIndex;	// Top index
			bool bNavigateLock = (Item.m_fsState & TCISTATE_NAVIGATELOCK) != 0;
			jsonTabData["NavigateLock"] = bNavigateLock;
			jsonTabData["ITEMIDLIST"] = ConvertBase64fromITEMIDLIST(GetItemIDList(i)).get();

			auto funcCreateJsonTravelLog = [](const TRAVELLOG& travelLog) -> json {
				json jsonTravelLog = jsonTravelLog.array();
				for (auto& tlItem : travelLog) {
					json jsTlItem;
					jsTlItem["title"] = CodeConvert::UTF8fromUTF16((LPCWSTR)tlItem.first);
					jsTlItem["URL"] = CodeConvert::UTF8fromUTF16((LPCWSTR)tlItem.second);
					jsonTravelLog.push_back(jsTlItem);
				}
				return jsonTravelLog;
			};
			jsonTabData["TravelLogBack"] = funcCreateJsonTravelLog(*Item.m_pTravelLogBack);
			jsonTabData["TravelLogFore"] = funcCreateJsonTravelLog(*Item.m_pTravelLogFore);
		}

		tabListArray.push_back(jsonTabData);
	}
	json jsonTabList;
	jsonTabList["TabList"] = tabListArray;
	
	CString	TabList = Misc::GetExeDirectory() + _T("TabList.json");
	CString	tempTabList = Misc::GetExeDirectory() + _T("TabList_temp.json");
	std::ofstream fs((LPCWSTR)tempTabList);
	//std::string dumpJson = jsonTabList.dump(4);
	fs << std::setw(4) << jsonTabList;
	fs.close();

	BOOL ret = ::MoveFileExW(tempTabList, TabList, MOVEFILE_REPLACE_EXISTING);
	ATLASSERT(ret);
}


void	CDonutTabBar::RestoreAllTab()
{
	CString	TabListJson = Misc::GetExeDirectory() + _T("TabList.json");
	if (::PathFileExists(TabListJson)) {
		try {
			std::ifstream fs((LPCWSTR)TabListJson);
			json jsonTabList;
			fs >> jsonTabList;
			fs.close();

			CComPtr<IMalloc> spMalloc;
			::SHGetMalloc(&spMalloc);
			ATLASSERT(spMalloc);

			json jsonTabListArray = jsonTabList["TabList"];
			for (auto& jsonTabItem : jsonTabListArray) {
				CTabItem item;
				item.m_strItem = CodeConvert::UTF16fromUTF8(jsonTabItem["Name"].get<std::string>()).c_str();
				if (item.m_strItem == L"<break>") {
					item.m_fsState |= TCISTATE_LINEBREAK;
				} else {
					item.m_strFullPath = CodeConvert::UTF16fromUTF8(jsonTabItem["FullPath"].get<std::string>()).c_str();
					item.m_nSelectedIndex = jsonTabItem["SelectedIndex"];
					bool bNavigateLock = jsonTabItem["NavigateLock"];
					if (bNavigateLock) {
						item.m_fsState |= TCISTATE_NAVIGATELOCK;
					}
					std::string base64PIDL = jsonTabItem["ITEMIDLIST"];
					int base64DecodeSize = ATL::Base64DecodeGetRequiredLength(static_cast<int>(base64PIDL.size()));
					LPITEMIDLIST  pNewIDList = (LPITEMIDLIST)spMalloc->Alloc(base64DecodeSize);
					::SecureZeroMemory(pNewIDList, base64DecodeSize);
					BOOL ret = ATL::Base64Decode(base64PIDL.c_str(), static_cast<int>(base64PIDL.size()), reinterpret_cast<BYTE*>(pNewIDList), &base64DecodeSize);
					ATLASSERT(ret);
					item.m_pidl = pNewIDList;

					auto funcCreateTravelLog = [](const json& jsonTravelLog) -> TRAVELLOG* {
						TRAVELLOG* travelLog = new TRAVELLOG;
						for (auto& jsTLItem : jsonTravelLog) {
							std::wstring title = CodeConvert::UTF16fromUTF8(jsTLItem["title"].get<std::string>());
							std::wstring URL = CodeConvert::UTF16fromUTF8(jsTLItem["URL"].get<std::string>());
							auto tlpair = std::make_pair<CString, CString>(title.c_str(), URL.c_str());
							travelLog->push_back(tlpair);
						}
						return travelLog;
						};
					item.m_pTravelLogBack = funcCreateTravelLog(jsonTabItem["TravelLogBack"]);
					item.m_pTravelLogFore = funcCreateTravelLog(jsonTabItem["TravelLogFore"]);
				}
				AddTabItem(item);
			}
		}
		catch (...) {
			ERROR_LOG << L"RestoreAllTab failed";
		}
	}
}

// ������ۑ�����
void	CDonutTabBar::SaveHistory()
{
	json historyListArray = json::array();
	for (const auto& historyItem : m_vecHistoryItem) {
		json jsonHistoryData;
		jsonHistoryData["Title"] = CodeConvert::UTF8fromUTF16((LPCWSTR)historyItem.strTitle);		
		jsonHistoryData["FullPath"] = CodeConvert::UTF8fromUTF16((LPCWSTR)historyItem.strFullPath);
		jsonHistoryData["ITEMIDLIST"] = ConvertBase64fromITEMIDLIST(historyItem.pidl).get();

		auto funcCreateJsonTravelLog = [](const TRAVELLOG& travelLog) -> json {
			json jsonTravelLog = jsonTravelLog.array();
			for (auto& tlItem : travelLog) {
				json jsTlItem;
				jsTlItem["title"] = CodeConvert::UTF8fromUTF16((LPCWSTR)tlItem.first);
				jsTlItem["URL"] = CodeConvert::UTF8fromUTF16((LPCWSTR)tlItem.second);
				jsonTravelLog.push_back(jsTlItem);
			}
			return jsonTravelLog;
		};
		jsonHistoryData["TravelLogBack"] = funcCreateJsonTravelLog(*historyItem.pLogBack);
		jsonHistoryData["TravelLogFore"] = funcCreateJsonTravelLog(*historyItem.pLogFore);

		historyListArray.push_back(jsonHistoryData);
	}
	json jsonHistoryList;
	jsonHistoryList["History"] = historyListArray;

	CString	Historyjson = Misc::GetExeDirectory() + _T("History.json");
	CString tempHistoryjson = Misc::GetExeDirectory() + _T("History_temp.json");
	std::ofstream fs((LPCWSTR)tempHistoryjson);
	//std::string dumpJson = jsonTabList.dump(4);
	fs << std::setw(4) << jsonHistoryList;
	fs.close();

	BOOL ret = ::MoveFileExW(tempHistoryjson, Historyjson, MOVEFILE_REPLACE_EXISTING);
	ATLASSERT(ret);
}

// �����𕜌�����
void	CDonutTabBar::RestoreHistory()
{
	CString	HistoryXML = Misc::GetExeDirectory() + _T("History.xml");
	if (::PathFileExists(HistoryXML)) {
		vector<HistoryItem>	vecItem;
		HistoryItem			item;

		vector<std::pair<int, BYTE*> > vecIDList;

		TRAVELLOG* pBackLog = NULL;
		TRAVELLOG* pForeLog = NULL;

		try {
			CXmlFileRead	xmlRead(HistoryXML);
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
										pBackLog->push_back(std::make_pair(title, url));
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
										pForeLog->push_back(std::make_pair(title, url));
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
						for (size_t i = 0; i < vecIDList.size(); ++i) {
							nSize += vecIDList[i].first;
						}

						LPITEMIDLIST pRetIDList;
						LPITEMIDLIST pNewIDList;

						CComPtr<IMalloc> spMalloc;
						::SHGetMalloc(&spMalloc);
						pNewIDList = pRetIDList = (LPITEMIDLIST)spMalloc->Alloc(nSize);
						::SecureZeroMemory(pRetIDList, nSize);

						std::size_t	nCount = vecIDList.size();
						for (size_t i = 0; i < nCount; ++i) {
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
			for (size_t i = 0; i < vecItem.size(); ++i) {
				HICON	hIcon = CreateIconFromIDList(vecItem[i].pidl);
				if (hIcon == NULL)
					hIcon = m_imgs.ExtractIcon(0);
				if (hIcon) {
					vecItem[i].hbmp = Misc::CreateBitmapFromHICON(hIcon);
					::DestroyIcon(hIcon);
				}
				m_vecHistoryItem.push_back(vecItem[i]);
			}

		} catch (LPCTSTR strError) {
			MessageBox(strError);
		} catch (...) {
			MessageBox(_T("RestoreHistory�Ɏ��s"));
		}

		SaveHistory();
		::DeleteFileW(HistoryXML);
		return;
	}

	CString	Historyjson = Misc::GetExeDirectory() + _T("History.json");
	if (::PathFileExists(Historyjson)) {
		std::ifstream fs((LPCWSTR)Historyjson);
		json jsonHistory;
		fs >> jsonHistory;
		fs.close();

		CComPtr<IMalloc> spMalloc;
		::SHGetMalloc(&spMalloc);
		ATLASSERT(spMalloc);

		json jsonHistoryArray = jsonHistory["History"];
		for (auto& jsonItem : jsonHistoryArray) {
			HistoryItem item;
			item.strTitle = CodeConvert::UTF16fromUTF8(jsonItem["Title"].get<std::string>()).c_str();
			item.strFullPath = CodeConvert::UTF16fromUTF8(jsonItem["FullPath"].get<std::string>()).c_str();			
			item.pidl = ConvertITEMIDLISTfromBase64(jsonItem["ITEMIDLIST"], spMalloc);

			auto funcCreateTravelLog = [](const json& jsonTravelLog) -> TRAVELLOG* {
				TRAVELLOG* travelLog = new TRAVELLOG;
				for (auto& jsTLItem : jsonTravelLog) {
					std::wstring title = CodeConvert::UTF16fromUTF8(jsTLItem["title"].get<std::string>());
					std::wstring URL = CodeConvert::UTF16fromUTF8(jsTLItem["URL"].get<std::string>());
					auto tlpair = std::make_pair<CString, CString>(title.c_str(), URL.c_str());
					travelLog->push_back(tlpair);
				}
				return travelLog;
			};
			item.pLogBack = funcCreateTravelLog(jsonItem["TravelLogBack"]);
			item.pLogFore = funcCreateTravelLog(jsonItem["TravelLogFore"]);

			HICON	hIcon = CreateIconFromIDList(item.pidl);
			if (hIcon == NULL)
				hIcon = m_imgs.ExtractIcon(0);
			if (hIcon) {
				item.hbmp = Misc::CreateBitmapFromHICON(hIcon);
				::DestroyIcon(hIcon);
			}
			m_vecHistoryItem.push_back(item);
		}
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
		if (_IsValidIndex(nNext) && (GetItem(nNext).m_fsState & TCISTATE_LINEBREAK)) {
			--nNext;
		}

		if (nNext >= 0) {
			return nNext;
		} else {
			nNext = nActiveIndex + 1;
			if (_IsValidIndex(nNext) && (GetItem(nNext).m_fsState & TCISTATE_LINEBREAK)) {
				++nNext;
			}
			if (nNext < nCount) {
				return nNext;
			}
		}
	} else {
		// ����Ƃ��A�N�e�B�u�ȃ^�u�̉E���A�N�e�B�u�ɂ���
		int nNext = nActiveIndex + 1;
		if (_IsValidIndex(nNext) && (GetItem(nNext).m_fsState & TCISTATE_LINEBREAK)) {
			++nNext;
		}
		if (nNext < nCount) {
			return nNext;
		} else {
			nNext = nActiveIndex - 1;
			if (_IsValidIndex(nNext) && (GetItem(nNext).m_fsState & TCISTATE_LINEBREAK)) {
				--nNext;
			}
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
	for (size_t i = 1; i < m_vecHistoryItem.size(); ++i) {
		if (m_vecHistoryItem[i].strFullPath == item.strFullPath) {
			_DeleteHistory(static_cast<int>(i));
			return;
		}
	}

	if (static_cast<int>(m_vecHistoryItem.size()) > CTabBarConfig::s_nMaxHistoryCount) {
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
	return;
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
#if 0
		const int itemCount = GetItemCount();
		for (int i = GetCurSel(); i < itemCount; ++i) {
			if (GetItem(i).m_fsState & TCISTATE_LINEBREAK) {
				nPos = i;
				break;
			}
		}
#endif
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
	SetItemText(nNewIndex, item.m_strItem);
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


void	CDonutTabBar::ExternalOpen(LPITEMIDLIST pidl, 
	boost::optional<OpenFolderAndSelectItems> folderAndSelectItems /*= boost::none*/)
{
	Misc::SetForegroundWindow(GetTopLevelWindow());
	std::thread([this]() { 
		::Sleep(50); 
		Misc::SetForegroundWindow(GetTopLevelWindow()); 
	}).detach();

	int nIndex = _IDListIsEqualIndex(pidl);
	if (nIndex != -1) {					
		if (nIndex != GetCurSel()) {		/* ���łɊJ���Ă�^�u������΂������Ɉړ����� */
			_SaveSelectedIndex(GetCurSel());
			SetCurSel(nIndex);
			m_folderAndSelectItems = folderAndSelectItems;

		} else {	// ���łɃA�N�e�B�u�ȃ^�u�ɊJ����Ă���
			if (folderAndSelectItems) {
				if (m_bTabChanged) {
					m_folderAndSelectItems = folderAndSelectItems;
					//INFO_LOG << L"ExternalOpen CurSelOpen tabchanged : true";

				} else {
					auto spFolderAndSelectItems = std::make_shared<OpenFolderAndSelectItems>(*folderAndSelectItems);
					std::thread([spFolderAndSelectItems]() {
						spFolderAndSelectItems->DoOpenFolderAndSelectItems();
					}).detach();
				}
			}
		}
		::CoTaskMemFree(pidl);

	} else {
		SetCurSel(OnTabCreate(pidl));
		m_folderAndSelectItems = folderAndSelectItems;
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

			{
				CMutexLock lock(m_mutex_currentPIDL);
				m_currentPIDL.CloseHandle();

				UINT ilSize = ::ILGetSize(pidl);
				void* view = m_currentPIDL.CreateSharedMemory(kCurrentPIDLSharedName, (DWORD)ilSize);
				::memcpy_s(view, ilSize, (void*)pidl, ilSize);
			}

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
	std::thread	thrd(std::bind(&CDonutTabBar::_threadVoidTabRemove, this));
	thrd.detach();
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
					::CoTaskMemFree(pidl);
					m_bDragItemIncludeFolder = true;
					break;
				}
				::CoTaskMemFree(pidl);
			}
			/* �h���b�O�A�C�e���� ���[�g�h���C�u(�V���[�g�J�b�g�������Ȃ�)���ǂ��� */
			m_bDragItemIsRoot = false;
			if (pida->cidl == 1) {
				LPCITEMIDLIST pChildIDList = GetPIDLItem(pida, 0);
				LPITEMIDLIST	pidl = ::ILCombine(pParentidl, pChildIDList);
				CString strDragItemPath = ShellWrap::GetFullPathFromIDList(pidl);
				::CoTaskMemFree(pidl);
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
		
		_hitTestFlag flag;
		int nIndex = HitTestOnDragging(flag, point, true);
		// �t�H���_�[���^�u���E��
		if (m_bDragItemIncludeFolder && (flag == _hitTestFlag::htSeparator || flag == _hitTestFlag::htInsetLeft)) {
			_DrawInsertionEdge(flag, nIndex);
			SetDropDescription(pDataObject, DROPIMAGE_LABEL, _T("�����Ƀ^�u���쐬"));
			return DROPEFFECT_LINK;
		}

		nIndex = HitTest(point);
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

		_hitTestFlag flag;
		int nInsertFolder = HitTestOnDragging(flag, point, true);
		// �t�H���_�[���^�u���E��
		if (!(m_bDragItemIncludeFolder && (flag == _hitTestFlag::htSeparator || flag == _hitTestFlag::htInsetLeft)))
			nInsertFolder = -1;

		HRESULT	hr;
		int nIndex = HitTest(point);
		if (nIndex == -1 || nInsertFolder != -1) {	// �^�u�O
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
						auto funcTabCreate = [&, this](LPITEMIDLIST pidl) {
							if (nInsertFolder != -1) {
								m_nInsertIndex = nInsertFolder;
								if (flag == _hitTestFlag::htSeparator)
									++m_nInsertIndex;
								OnTabCreate(pidl, false, true);
								m_nInsertIndex = -1;
								++nInsertFolder;
							} else {
								OnTabCreate(pidl, true);
							}
						};
						SFGAOF	attribute;
						spShellItem->GetAttributes(SFGAO_FOLDER | SFGAO_LINK, &attribute);
						if (attribute & SFGAO_LINK) {	// lnk�Ȃ��������
							LPITEMIDLIST	pidlLink = ShellWrap::GetResolveIDList(pidl);
							if (pidlLink) {
								funcTabCreate(pidlLink);
							} else if (attribute & SFGAO_FOLDER) {
								funcTabCreate(pidl);
							}
						} else if (attribute & SFGAO_FOLDER) {
							funcTabCreate(pidl);
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
			std::thread	thrd(std::bind(&CDonutTabBar::_threadPerformSHFileOperation, this, GetItemIDList(nIndex), pDataObject, dropEffect == DROPEFFECT_MOVE));
			thrd.detach();
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
	INFO_LOG << L"RefreshTab: " << title;

	LPITEMIDLIST	pidl = GetCurIDList(m_spShellBrowser);
	if (pidl == NULL) {
		/* �^�u����ύX */
		if (GetCurSel() != -1)
			SetItemText(GetCurSel(), title);	// Windows\assembly �t�H���_�Ƃ��ŕK�v
		return;		// ���݂̃t�H���_�̃A�C�e���h�c���X�g���擾�ł��Ȃ�����
	}

	int nCurIndex = GetCurSel();
	if (nCurIndex == -1) {	// �^�u������Ȃ��Ȃ�G�N�X�v���[���[���N�������Ƃ���
		CString strFullPath = GetFullPathFromIDList(pidl);
		if (   strFullPath.Left(5)  == _T("��������&")
			|| (CTabBarConfig::s_bMargeControlPanel == false 
				&& strFullPath.Left(40) == _T("::{26EE0668-A00A-44D7-9371-BEB064C98683}"))) 
		{
			m_bSaveAllTab = false;
			_RefreshBandInfo();
			::CoTaskMemFree(pidl);
			return ;	// �������ʕ\���͕ʃE�B���h�E�ɂ���
		}
		HWND hWndTarget = FindWindow(_T("ExpTabBar_NotifyWindow"), NULL);
		if (hWndTarget != NULL) {	/* ���ɂ��G�N�X�v���[���[���N�����Ă��� */			
			m_bSaveAllTab = false;

			// �O������t�H���_���J���ꂽ��
			if (CNotifyWindow::GetInstance().GetTabBarCount() == 1) {
				INFO_LOG << L"RefreshTab: Notify another explorer [" << (LPCWSTR)strFullPath << L"]";

				// ���̃G�N�X�v���[���[��pidl���J���ꂽ���Ƃ�ʒm����
				COPYDATASTRUCT	cd = { 0 };
				UINT cbItemID = ::ILGetSize(pidl);
				cd.lpData = (LPVOID)pidl;
				cd.cbData = cbItemID;
				SendMessage(hWndTarget, WM_COPYDATA, NULL, (LPARAM)&cd);
				::CoTaskMemFree(pidl);

				::PostMessage(GetTopLevelWindow(), WM_CLOSE, 0, 0);


				std::thread([] {
					::Sleep(5 * 1000);
					INFO_LOG << L"Terminate Explorer!";

					DWORD processID = GetCurrentProcessId();
					HANDLE h = ::OpenProcess(PROCESS_TERMINATE, FALSE, processID);
					ATLASSERT(h);
					::TerminateProcess(h, 0);
					::CloseHandle(h);
				}).detach();

				return;
			} else {
				if (CNotifyWindow::GetInstance().m_hWnd == NULL) {
					// �ʒm���󂯎�邽�߂̃E�B���h�E�����
					INFO_LOG << L"CNotifyWindow::GetInstance().CreateEx(NULL); - ���v���Z�X�� ExpTabBar_NotifyWindow �����݂���";
					CNotifyWindow::GetInstance().CreateEx(NULL);
				}

				// �����͂ǂ������V�i���I�Ȃ񂾁H
				// �ʃv���Z�X�̃G�N�X�v���[���[�ƒʒm�E�B���h�E�̃n���h�����قȂ�Ƃ�

				INFO_LOG << L"TopTabCreate [" << (LPCWSTR)strFullPath << L"]";
				// �g�b�v�^�u���ǉ����ꂽ���H
				// �V�K�^�u�쐬������
				nCurIndex = OnTabCreate(pidl, true);
				SetCurSel(nCurIndex);
				return;
			}




			return;
#if 0
			OSVERSIONINFO	osvi = { sizeof(osvi) };
			GetVersionEx(&osvi);
			if (osvi.dwMajorVersion == 6 && osvi.dwMinorVersion == 1) {	// Win7
				std::thread([]{
					::Sleep(5 * 1000);

					DWORD processID = GetCurrentProcessId();
					HANDLE h = ::OpenProcess(PROCESS_TERMINATE, FALSE, processID);
					ATLASSERT(h);
					::TerminateProcess(h, 0);
					::CloseHandle(h);
				}).detach();
			}
#endif
			return;
		} else {	// ���̃G�N�X�v���[���[�͂Ȃ�

			// �ʒm���󂯎�邽�߂̃E�B���h�E�����
			INFO_LOG << L"CNotifyWindow::GetInstance().CreateEx(NULL);";
			CNotifyWindow::GetInstance().CreateEx(NULL);

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
		if ((GetItemState(nCurIndex) & TCISTATE_NAVIGATELOCK) == 0)
			_AddHistory(nCurIndex);	// �ŋߕ����^�u�ɒǉ�����

		SetCurSel(nIndex);

		if ((GetItemState(nCurIndex) & TCISTATE_NAVIGATELOCK) == 0)
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

	/* �t���p�X��ύX */
	SetItemFullPath(nCurIndex, strFullPath);

	/* �^�u����ύX */
	SetItemText(nCurIndex, title/*GetNameFromIDList(pidl)*/);

	/* �A�C�R����ύX */
	//CreateIconFromIDList(pidl);
	ReplaceIcon(nCurIndex, CreateIconFromIDList(pidl));

	/* �A�C�e���h�c���X�g��ύX */
	SetItemIDList(nCurIndex, pidl);

	{
		CMutexLock lock(m_mutex_currentPIDL);
		m_currentPIDL.CloseHandle();

		UINT ilSize = ::ILGetSize(pidl);
		void* view = m_currentPIDL.CreateSharedMemory(kCurrentPIDLSharedName, (DWORD)ilSize);
		::memcpy_s(view, ilSize, (void*)pidl, ilSize);
	}
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

		//_RestoreSelectedIndex();
		m_restoreSelectedIndexRetryCount = 0;
		SetTimer(RestoreScrollPosTimerID, 100);
		
		if (m_folderAndSelectItems) {
			int nIndex = _IDListIsEqualIndex((LPITEMIDLIST)m_folderAndSelectItems->pidlFolder.data());
			if (nIndex == GetCurSel()) {
				m_folderAndSelectItems->DoOpenFolderAndSelectItems();
			}
			m_folderAndSelectItems.reset();
		}
	}
#endif
}

void CDonutTabBar::TopTabActivate()
{
	CNotifyWindow::GetInstance().ActiveTabBar(this);
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
	//ATLASSERT(m_hSearch);

	m_menuPopup.LoadMenu(IDM_TAB);
	m_menuHistory = m_menuPopup.GetSubMenu(1).GetSubMenu(0);
	m_menuFavorites = m_menuPopup.GetSubMenu(1).GetSubMenu(1);

	SetTimer(AutoSaveTimerID, AutoSaveInterval);

	CNotifyWindow::GetInstance().AddTabBar(this);

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
		m_bSaveAllTab = false;
	}

	m_tipHistroy.DestroyWindow();

	m_menuPopup.DestroyMenu();

	KillTimer(AutoSaveTimerID);

	CNotifyWindow::GetInstance().RemoveTabBar(this);

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
		
		CMenuHandle menu = m_menuPopup.GetSubMenu(1);

		int	i = CTabBarConfig::s_bMultiLine ? 0 : m_nFirstIndexOnSingleLine;
		int nCount = GetItemCount();
		int offsetY = 0;
		m_tabGroupIndexOnRClick = -1;
		CRect rcClient;
		GetClientRect(&rcClient);
		for (; i < nCount; ++i) {
			if (m_items[i].m_fsState & TCISTATE_LINEBREAK) {
				++m_tabGroupIndexOnRClick;

				CRect rcGroup;
				rcGroup.top = offsetY;
				rcGroup.left = 0;
				rcGroup.right = rcClient.right;
				rcGroup.bottom = m_items[i].m_rcItem.bottom;
				offsetY = rcGroup.bottom;
				if (rcGroup.PtInRect(point)) {
					break;
				}
			} else if ((i + 1) == nCount && m_tabGroupIndexOnRClick >= 0) {
				++m_tabGroupIndexOnRClick;
			}
		}
		enum { kTabGroupMenuIndex = 3 };
		if (m_tabGroupIndexOnRClick == -1) {
			menu.EnableMenuItem(kTabGroupMenuIndex, MF_BYPOSITION | MF_GRAYED);
		} else {
			menu.EnableMenuItem(kTabGroupMenuIndex, MF_BYPOSITION | MF_ENABLED);
		}

		m_tipHistroy.SetWindowPos(HWND_TOPMOST, -1, -1, -1, -1, SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE);

		CToolInfo	tinfo(TTF_SUBCLASS | TTF_TRACK, m_hWnd, 0, NULL, _T(""));
		m_tipHistroy.AddTool(tinfo);
		m_tipHistroy.Popup();

		m_tipHistroy.Activate(TRUE);
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
			if (i < 0)
				return;
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
			if (i < 0)
				return ;
			const CString& strPath = CFavoritesOption::s_vecFavoritesItem[i].strPath;
			if (strPath == FAVORITESSEPSTRING)
				return ;
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
	} else if (nIDEvent == RestoreScrollPosTimerID) {
		KillTimer(RestoreScrollPosTimerID);

		enum { kMaxRetryCount = 20  };
		if (m_restoreSelectedIndexRetryCount < kMaxRetryCount) {
			++m_restoreSelectedIndexRetryCount;
			if (!_RestoreSelectedIndex()) {
				SetTimer(RestoreScrollPosTimerID, 100);
			}
		}
		//_RestoreSelectedIndex();
	} else {
		SetMsgHandled(FALSE);
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
		if (zDelta > 0) {	// �O�̃^�u
			nNext  = nIndex - 1 < 0 ? nCount - 1 : nIndex - 1;
			if (GetItem(nNext).m_fsState & TCISTATE_LINEBREAK) {
				--nIndex;
				nNext = nIndex - 1 < 0 ? nCount - 1 : nIndex - 1;
			}
		} else {	// ���̃^�u
			nNext  = (nIndex + 1 < nCount) ? nIndex + 1 : 0;
			if (GetItem(nNext).m_fsState & TCISTATE_LINEBREAK) {
				++nIndex;
				nNext = (nIndex + 1 < nCount) ? nIndex + 1 : 0;
			}
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

		/* ���łɊJ���Ă�^�u������΂������Ɉړ����� */
		int nAlreadyIndex = _IDListIsEqualIndex(pidl);
		if (nAlreadyIndex != -1 && nAlreadyIndex != nIndex) {
			_SaveSelectedIndex(nIndex);

			SetCurSel(nAlreadyIndex);

		} else {
			m_nInsertIndex = nIndex;
			int nNewIndex = OnTabCreate(pidl, false, true);
			SetCurSel(nNewIndex);
			m_nInsertIndex = -1;
		}
	}
}

/// ���̃^�u����^�u�O���[�v���쐬
void	CDonutTabBar::OnCreateTabGroup(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	if (m_ClickedIndex == -1 || GetItemCount() <= 1) {
		return;
	}

	CSimpleArray<int> arrCurMultiSel;
	GetCurMultiSelEx(arrCurMultiSel, m_ClickedIndex);

	MoveItems(0, arrCurMultiSel);

	CTabItem item(L"<break>");
	item.ModifyState(0, TCISTATE_LINEBREAK);
	InsertItem(arrCurMultiSel.GetSize(), item);

	_CorrectTabGroupBorder();
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
	bool bOldShowParentFoldreName = CTabBarConfig::s_bShowParentFolderNameIfSameName;

	CExpTabBarOption	dlg;
	dlg.Show(m_hWnd);

	if (bOldShowParentFoldreName != CTabBarConfig::s_bShowParentFolderNameIfSameName) {
		const int itemCount = GetItemCount();
		for (int i = 0; i < itemCount; ++i) {
			SetItemText(i, GetItemText(i));
		}
	}

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
	//AddTabItem(item);
	int insertIndex = -1;
	const int itemCount = GetItemCount();
	for (int i = GetCurSel(); i < itemCount; ++i) {
		if (GetItem(i).m_fsState & TCISTATE_LINEBREAK) {
			insertIndex = i;
			break;
		}
	}
	HICON hIcon = CreateIconFromIDList(item.m_pidl);
	if (hIcon) {
		item.m_nImgIndex = AddIcon(hIcon);
	} else {
		item.m_nImgIndex = 0;
	}
	InsertItem(insertIndex, item);

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

void CDonutTabBar::OnMoveTabGroup(UINT uNotifyCode, int nID, CWindow wndCtl)
{
	ATLASSERT(m_tabGroupIndexOnRClick != -1);

	bool bMoveUp = ID_MOVEUP_TABGROUP == nID;
	CSimpleArray<int> arrTabGroupIndexs;

	int	i = CTabBarConfig::s_bMultiLine ? 0 : m_nFirstIndexOnSingleLine;
	int nCount = GetItemCount();
	int tabGroupIndex = 0;
	int upGroupFrontIndex = -1;
	int downGroupBottomIndex = -1;
	int& insertPos = bMoveUp ? upGroupFrontIndex : downGroupBottomIndex;
	for (; i < nCount; ++i) {
		if (tabGroupIndex == m_tabGroupIndexOnRClick) {
			arrTabGroupIndexs.Add(i);
		} else if ((m_tabGroupIndexOnRClick - 1) == tabGroupIndex) {
			if (upGroupFrontIndex == -1) {
				upGroupFrontIndex = i;
			}
		} else if ((m_tabGroupIndexOnRClick + 1) == tabGroupIndex) {
			downGroupBottomIndex = i + 1;
		}

		if (m_items[i].m_fsState & TCISTATE_LINEBREAK) {
			++tabGroupIndex;
		} else if ((i + 1) == nCount) {	// �Ō�̃^�u�̎�
			if (tabGroupIndex == m_tabGroupIndexOnRClick) {	// �P�c�̃^�u�O���[�v����Ɉړ�������Ƃ�
				// �Ō�̃^�u�O���[�v�ɂ̓{�[�_�[���Ȃ��̂Œǉ����Ă���
				if (insertPos != -1) {
					CTabItem item(L"<break>");
					item.ModifyState(0, TCISTATE_LINEBREAK);
					InsertItem(nCount, item);

					arrTabGroupIndexs.Add(nCount);
				}
			} else if ((m_tabGroupIndexOnRClick + 1) == tabGroupIndex) {	// �P�c�����O�̃O���[�v�����Ɉړ�������Ƃ�)
				// �Ō�̃^�u�O���[�v�ɂ̓{�[�_�[���Ȃ��̂Œǉ����Ă���
				if (!bMoveUp) {
					CTabItem item(L"<break>");
					item.ModifyState(0, TCISTATE_LINEBREAK);
					InsertItem(nCount, item);

					++downGroupBottomIndex;
				}
			}
		}
	}
	ATLASSERT(arrTabGroupIndexs.GetSize() > 0);
	if (insertPos != -1) {
		MoveItems(insertPos, arrTabGroupIndexs);
	}	

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
			CComQIPtr<IFolderView2> spFolderView = spShellView;
			if (spFolderView) {
				UINT viewMode = 0;
				spFolderView->GetCurrentViewMode(&viewMode);
				ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
				if (ListView == NULL) {
					HWND hwndUIItemsView = ::GetWindow(hWnd, GW_CHILD);
					if (hwndUIItemsView) {
						//double scrollPercent = m_UIAWrapper.GetUIItemsViewVerticalScrollPercent(hwndUIItemsView);
						// INT_MAX -> 2147483647 : 10��
						// double 100.0 * 10000000 = 1000000000 : 10��
						//int scrollPercentInt = static_cast<int>(scrollPercent * 10000000);
						//SetItemSelectedIndex(nIndex, scrollPercentInt);
					}
				} else {
					if (viewMode == FVM_ICON) {
						int scrollPos2 = m_UIAWrapper.GetScrollPos(ListView.m_hWnd);
						SetItemSelectedIndex(nIndex, scrollPos2);
					} else {
						SetItemSelectedIndex(nIndex, ListView.GetTopIndex());
					}
				}
			}
		}
	}
}

bool CDonutTabBar::_RestoreSelectedIndex()
{
	CListViewCtrl ListView;
	HWND hWnd;
	CComPtr<IShellView>	spShellView;
	HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
	if (SUCCEEDED(hr)) {
		hr = spShellView->GetWindow(&hWnd);
		if (SUCCEEDED(hr)) {
			CComQIPtr<IFolderView2>	spFolderView = spShellView;
			if (spFolderView) {
				UINT viewMode = 0;
				spFolderView->GetCurrentViewMode(&viewMode);
				ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
				if (ListView == NULL) {
					HWND hwndUIItemsView = ::GetWindow(hWnd, GW_CHILD);
					if (hwndUIItemsView) {
						//double selectedIndex = GetItemSelectedIndex(GetCurSel());
						//double scrollPercent = selectedIndex / 10000000;
						//m_UIAWrapper.SetUIItemsViewVerticalScrollPercent(hwndUIItemsView, scrollPercent);
					}
				} else {
					if (viewMode == FVM_ICON) {
						if (!m_UIAWrapper.SetScrollPos(ListView, GetItemSelectedIndex(GetCurSel()))) {
							return false;
						}
					} else {
						int nPerPage = ListView.GetCountPerPage();
						ListView.EnsureVisible(GetItemSelectedIndex(GetCurSel()) + nPerPage - 1, FALSE);
					}
				} 
			}
		}
	}
	return true;
}




void	CDonutTabBar::_threadPerformSHFileOperation(LPITEMIDLIST pidlTo, IDataObject* pDataObject, bool bMove)
{
	::CoInitialize(NULL);

	auto vecIDList = ShellWrap::GetIDListFromDataObject(pDataObject);
	bool bIsLink = false;
	if (vecIDList.size() == 1) {
		CString fullPath = ShellWrap::GetFullPathFromIDList(vecIDList[0]);
		if (fullPath.Left(2) == _T("::") || (fullPath.GetLength() == 3 && fullPath.Mid(1, 2) == _T(":\\")))
			bIsLink = true;
	}

	// �����N���쐬����
	if ( (::GetKeyState(VK_SHIFT) < 0 && ::GetKeyState(VK_CONTROL) < 0) || bIsLink ) {
		CString strTargetFolder = ShellWrap::GetFullPathFromIDList(pidlTo);
		for (auto it = vecIDList.begin(); it != vecIDList.end(); ++it) {
			LPITEMIDLIST pidl = *it;
			CString strLinkName = ShellWrap::GetNameFromIDList(pidl);
			strLinkName.Replace(L':', L'�F');
			CString strLinkPath;
			strLinkPath.Format(_T("%s\\%s.lnk"), (LPCWSTR)strTargetFolder, (LPCWSTR)strLinkName);
			ShellWrap::CreateLinkFile(pidl, strLinkPath);
			::CoTaskMemFree(pidl);
		}
	} else {
		std::for_each(vecIDList.begin(), vecIDList.end(), [](LPITEMIDLIST pidl) {
			::CoTaskMemFree(pidl);
		});

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













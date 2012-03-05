/**
*	@file	ExpBabBarOption.cpp
*	@brief	オプション
*/

#include "stdafx.h"
#include "ExpTabBarOption.h"
#include "resource.h"
#include <codecvt>
#include <atlddx.h>
#include <boost/property_tree/ptree.hpp>
#include <boost/property_tree/ini_parser.hpp>
#include <boost/property_tree/xml_parser.hpp>
#include <boost/optional.hpp>
#include <atlenc.h>
#include "DonutFunc.h"
#include "ShellWrap.h"

using namespace boost::property_tree;

/////////////////////////////////////////
// タブの設定

// 定義
int		CTabBarConfig::s_nAddPos			= RIGHTPOS;
bool	CTabBarConfig::s_bLeftActiveOnClose	= true;
bool	CTabBarConfig::s_bAddLinkRight		= true;
bool	CTabBarConfig::s_bLinkActive		= false;
bool	CTabBarConfig::s_bWheel				= false;
bool	CTabBarConfig::s_bMultiLine			= true;
bool	CTabBarConfig::s_bUseFixedSize		= false;
CSize	CTabBarConfig::s_FixedSize(100, 24);
int		CTabBarConfig::s_nMaxTextLength		= 30;
int		CTabBarConfig::s_RClickCommand		= SHOWMENU;
int		CTabBarConfig::s_DblClickCommand	= TABCLOSE;
int		CTabBarConfig::s_MClickCommand		= TABCLOSE;
int		CTabBarConfig::s_nMaxHistoryCount	= 16;
bool	CTabBarConfig::s_bMargeControlPanel = false;
bool	CTabBarConfig::s_bNoFullRowSelect	= true;

/// 設定を読み込む
void	CTabBarConfig::LoadConfig()
{
	ptree pt;
	std::ifstream	inistream(g_szIniFileName);
	read_ini(inistream, pt);

	if (auto value = pt.get_optional<int>("Tab.AddPos"))
		s_nAddPos = value.get();
	if (auto value = pt.get_optional<bool>("Tab.LeftActiveOnClose"))
		s_bLeftActiveOnClose = value.get();
	if (auto value = pt.get_optional<bool>("Tab.AddLinkRight"))
		s_bAddLinkRight = value.get();
	if (auto value = pt.get_optional<bool>("Tab.LinkActive"))
		s_bLinkActive	= value.get();
	if (auto value = pt.get_optional<bool>("Tab.Wheel"))
		s_bWheel	= value.get();
	if (auto value = pt.get_optional<bool>("Tab.MultiLine"))
		s_bMultiLine	= value.get();
	if (auto value = pt.get_optional<bool>("Tab.UseFixedSize"))
		s_bUseFixedSize	= value.get();
	if (auto value = pt.get_optional<int>("Tab.FixedSizeX"))
		s_FixedSize.cx	= value.get();
	if (auto value = pt.get_optional<int>("Tab.FixedSizeY"))
		s_FixedSize.cy	= value.get();
	if (auto value = pt.get_optional<int>("Tab.MaxTextLength"))
		s_nMaxTextLength	= value.get();
	if (auto value = pt.get_optional<int>("Tab.RClickCommand"))
		s_RClickCommand		= value.get();
	if (auto value = pt.get_optional<int>("Tab.DblClickCommand"))
		s_DblClickCommand	= value.get();
	if (auto value = pt.get_optional<int>("Tab.MClickCommand"))
		s_MClickCommand		= value.get();
	if (auto value = pt.get_optional<int>("Tab.MaxHistoryCount"))
		s_nMaxHistoryCount	= value.get();
	if (auto value = pt.get_optional<bool>("Tab.MargeControlPanel"))
		s_bMargeControlPanel = value.get();
	if (auto value = pt.get_optional<bool>("Tab.NoFullRowSelect"))
		s_bNoFullRowSelect	= value.get();
}

/// 設定を保存する
void	CTabBarConfig::SaveConfig()
{
	ptree pt;
	std::ifstream	inistream(g_szIniFileName);
	read_ini(inistream, pt);
	inistream.close();

	pt.put("Tab.AddPos"				, s_nAddPos);
	pt.put("Tab.LeftActiveOnClose"	, s_bLeftActiveOnClose);
	pt.put("Tab.AddLinkRight"		, s_bAddLinkRight);
	pt.put("Tab.LinkActive"			, s_bLinkActive);
	pt.put("Tab.Wheel"				, s_bWheel);
	pt.put("Tab.MultiLine"			, s_bMultiLine);
	pt.put("Tab.UseFixedSize"		, s_bUseFixedSize);
	pt.put("Tab.FixedSizeX"			, s_FixedSize.cx);
	pt.put("Tab.FixedSizeY"			, s_FixedSize.cy);
	pt.put("Tab.MaxTextLength"		, s_nMaxTextLength);
	pt.put("Tab.RClickCommand"		, s_RClickCommand);
	pt.put("Tab.DblClickCommand"	, s_DblClickCommand);
	pt.put("Tab.MClickCommand"		, s_MClickCommand);
	pt.put("Tab.MaxHistoryCount"	, s_nMaxHistoryCount);
	pt.put("Tab.MargeControlPanel"	, s_bMargeControlPanel);
	pt.put("Tab.NoFullRowSelect"	, s_bNoFullRowSelect);

	std::ofstream iniostream(g_szIniFileName, std::ios::out | std::ios::trunc);
	write_ini(iniostream, pt);
	iniostream.close();
}




//////////////////////////////////////////////
/// タブ
class CTabBarOptionPropertyPage : 
	public CPropertyPageImpl<CTabBarOptionPropertyPage>,
	public CWinDataExchange<CTabBarOptionPropertyPage>,
	protected CTabBarConfig
{
public:
	enum { IDD = IDD_TABBAROPTION };

	// DDX map
    BEGIN_DDX_MAP(CTabBarOptionPropertyPage)
		DDX_CONTROL_HANDLE(IDC_CMB_ADDPOS		, m_cmbAddPos )
		DDX_CONTROL_HANDLE(IDC_CMB_POSONCLOSE	, m_cmbPosOnClose )
		DDX_CHECK(IDC_CHECK_ADDLINKRIGHT		, s_bAddLinkRight )
		DDX_CHECK(IDC_CHECK_LINKOPENACTIVE		, s_bLinkActive )

		DDX_CONTROL_HANDLE(IDC_CMB_RCLICK		, m_cmbRClick )
		DDX_CONTROL_HANDLE(IDC_CMB_DBLCLICK		, m_cmbDblClick )
		DDX_CONTROL_HANDLE(IDC_CMB_MCLICK		, m_cmbMClick )
		DDX_CHECK(IDC_CHECK_USEWHEEL			, s_bWheel )

		DDX_CHECK(IDC_CHECK_MULTILINE			, s_bMultiLine )
		DDX_CHECK(IDC_CHECK_FIXEDSIZE			, s_bUseFixedSize )
		DDX_INT_RANGE(IDC_EDIT_FIXED_X			, (int&)s_FixedSize.cx	, 10, 500)
		DDX_INT_RANGE(IDC_EDIT_FIXED_Y			, (int&)s_FixedSize.cy	, 10, 500)
		DDX_INT_RANGE(IDC_EDIT_MAXTEXTLENGTH	, s_nMaxTextLength	, 10, 125)
		DDX_INT_RANGE(IDC_EDIT_MAXHISTORYCOUNT	, s_nMaxHistoryCount,  1,  50)
		DDX_CHECK(IDC_CHECK_MARGECONTROLPANEL	, s_bMargeControlPanel	)
		DDX_CHECK(IDC_CHECK_NOFULLROWSELECT		, s_bNoFullRowSelect	)
    END_DDX_MAP()
	

	// Message map
	BEGIN_MSG_MAP_EX( CTabBarOptionPropertyPage )
		MSG_WM_INITDIALOG( OnInitDialog )
		CHAIN_MSG_MAP( CPropertyPageImpl<CTabBarOptionPropertyPage> )
	END_MSG_MAP()

	BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam)
	{
		DoDataExchange(DDX_LOAD);

		/* 追加位置 */
		m_cmbAddPos.AddString(_T("右端"));
		m_cmbAddPos.AddString(_T("左端"));
		m_cmbAddPos.AddString(_T("アクティブの右側"));
		m_cmbAddPos.AddString(_T("アクティブの左側"));
		m_cmbAddPos.SetCurSel(s_nAddPos);

		/* 閉じたときの位置 */
		m_cmbPosOnClose.AddString(_T("左側をアクティブにする"));
		m_cmbPosOnClose.AddString(_T("右側をアクティブにする"));
		m_cmbPosOnClose.SetCurSel(s_bLeftActiveOnClose ? 0 : 1);


		struct ComboItem {
			LPCTSTR	strItem;
			UINT	nID;
		};

		ComboItem items[] = {
			{ _T("なし")					, COMMANDNONE	},
			{ _T("閉じる")					, TABCLOSE	},
			{ _T("一つ上のフォルダを開く")	, OPEN_UPFOLDER	},
			{ _T("ナビゲートロックする")	, NAVIGATELOCK	},
			{ _T("メニューを表示する")		, SHOWMENU	},
		};
		for (int i = 0; i < _countof(items); ++i) {
			m_cmbRClick.AddString(items[i].strItem);
			m_cmbDblClick.AddString(items[i].strItem);
			m_cmbMClick.AddString(items[i].strItem);
		}
		m_cmbRClick.SetCurSel(s_RClickCommand);
		m_cmbDblClick.SetCurSel(s_DblClickCommand);
		m_cmbMClick.SetCurSel(s_MClickCommand);

		/* 複数列タブ以外は今のところ無理 */
		GetDlgItem(IDC_CHECK_MULTILINE).EnableWindow(FALSE);
		return 0;
	}

	BOOL OnApply()
	{
		if (!DoDataExchange(DDX_SAVE))
			return FALSE;

		s_nAddPos	= m_cmbAddPos.GetCurSel();
		s_bLeftActiveOnClose	= m_cmbPosOnClose.GetCurSel() == 0;

		s_RClickCommand		= m_cmbRClick.GetCurSel();
		s_DblClickCommand	= m_cmbDblClick.GetCurSel();
		s_MClickCommand		= m_cmbMClick.GetCurSel();
		
		SaveConfig();	/* 保存 */

		return TRUE;
	}

private:
	// Data members
	CComboBox	m_cmbAddPos;
	CComboBox	m_cmbPosOnClose;
	CComboBox	m_cmbRClick;
	CComboBox	m_cmbDblClick;
	CComboBox	m_cmbMClick;
};

//////////////////////////////////////////
// サムネイルツールチップの設定

// 定義
bool	CThumbnailTooltipConfig::s_bUseThumbnailTooltip = false;
CSize	CThumbnailTooltipConfig::s_MaxThumbnailSize(512, 256);
int		CThumbnailTooltipConfig::s_nMaxThumbnailCache = 64;

bool	CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged = false;

void CThumbnailTooltipConfig::LoadConfig()
{
	ptree pt;
	std::ifstream	inistream(g_szIniFileName);
	read_ini(inistream, pt);

	if (auto value = pt.get_optional<bool>("Thumbtip.UseThumbnailTooltip"))
		s_bUseThumbnailTooltip = value.get();
	if (auto value = pt.get_optional<int>("Thumbtip.MaxThumbnailSizeWidth"))
		s_MaxThumbnailSize.cx	= value.get();
	if (auto value = pt.get_optional<int>("Thumbtip.MaxThumbnailSizeHeight"))
		s_MaxThumbnailSize.cy	= value.get();
	if (auto value = pt.get_optional<int>("Thumbtip.MaxThumbnailCache"))
		s_nMaxThumbnailCache	= value.get();
}

void CThumbnailTooltipConfig::SaveConfig()
{
	ptree pt;
	std::ifstream	inistream(g_szIniFileName);
	read_ini(inistream, pt);
	inistream.close();

	pt.put("Thumbtip.UseThumbnailTooltip"	, s_bUseThumbnailTooltip);
	pt.put("Thumbtip.MaxThumbnailSizeWidth"	, s_MaxThumbnailSize.cx);
	pt.put("Thumbtip.MaxThumbnailSizeHeight", s_MaxThumbnailSize.cy);
	pt.put("Thumbtip.MaxThumbnailCache"		, s_nMaxThumbnailCache);

	std::ofstream iniostream(g_szIniFileName, std::ios::out | std::ios::trunc);
	write_ini(iniostream, pt);
	iniostream.close();
}


//////////////////////////////////////////////
/// サムネイルツールチップ
class CThumbnailTooltipPropertyPage : 
	public CPropertyPageImpl<CThumbnailTooltipPropertyPage>,
	public CWinDataExchange<CThumbnailTooltipPropertyPage>,
	protected CThumbnailTooltipConfig
{
public:
	enum { IDD = IDD_THUMBNAILTOOLTIP };

	// DDX map
    BEGIN_DDX_MAP(CThumbnailTooltipPropertyPage)
		DDX_CHECK(IDC_CHECK_USRTHUMBNAILTOOLTIP	, s_bUseThumbnailTooltip )
		DDX_INT_RANGE(IDC_EDIT_MAXWIDTH			, (int&)s_MaxThumbnailSize.cx	, 64, 1024)
		DDX_INT_RANGE(IDC_EDIT_MAXHEIGHT		, (int&)s_MaxThumbnailSize.cy	, 64, 1024)
		DDX_INT_RANGE(IDC_EDIT_MAXTHUMBNAIL		, s_nMaxThumbnailCache	, 1, 2000)
    END_DDX_MAP()
	

	// Message map
	BEGIN_MSG_MAP_EX( CThumbnailTooltipPropertyPage )
		MSG_WM_INITDIALOG( OnInitDialog )
		CHAIN_MSG_MAP( CPropertyPageImpl<CThumbnailTooltipPropertyPage> )
	END_MSG_MAP()

	BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam)
	{
		DoDataExchange(DDX_LOAD);
		CUpDownCtrl(GetDlgItem(IDC_SPIN1)).SetRange(64, 1024);
		CUpDownCtrl(GetDlgItem(IDC_SPIN2)).SetRange(64, 1024);
		CUpDownCtrl(GetDlgItem(IDC_SPIN3)).SetRange(1, 2000);

		return 0;
	}

	BOOL OnApply()
	{
		CSize prevSize = s_MaxThumbnailSize;
		if (!DoDataExchange(DDX_SAVE))
			return FALSE;

		if (prevSize != s_MaxThumbnailSize)
			s_bMaxThumbnailSizeChanged	= true;	// キャッシュのクリアを促す

		SaveConfig();	/* 保存 */

		return TRUE;
	}

    void OnDataValidateError(UINT nCtrlID, BOOL bSave, _XData& data)
	{
        CString strMsg; 
        strMsg.Format(_T("%d から %d までの値を入力してください。"),
            data.intData.nMin, data.intData.nMax);
        MessageBox(strMsg, _T("エラー"), MB_ICONEXCLAMATION);

        ::SetFocus(GetDlgItem(nCtrlID));
	}

};

////////////////////////////////////////////
// お気に入りの設定

// 定義
std::vector<CFavoritesOption::FavoritesItem>	CFavoritesOption::s_vecFavoritesItem;

/// pidl をお気に入りに追加する
void CFavoritesOption::AddFavorites(LPITEMIDLIST pidl)
{
	ATLASSERT(pidl);
	FavoritesItem	item;
	item.strPath	= ShellWrap::GetFullPathFromIDList(pidl);
	item.strTitle	= ShellWrap::GetNameFromIDList(pidl);
	item.pidl		= ::ILClone(pidl);
	CIcon	icon = ShellWrap::CreateIconFromIDList(pidl);
	if (icon.m_hIcon) 
		item.bmpIcon	= Misc::CreateBitmapFromHICON(icon);
	s_vecFavoritesItem.push_back(item);

	SaveConfig();
}

/// s_vecFavoritesItem をかたずける
void	CFavoritesOption::CleanFavoritesItem()
{
	for (auto it = s_vecFavoritesItem.begin(); it != s_vecFavoritesItem.end(); ++it) {
		if (it->pidl)
			::ILFree(it->pidl);
		if (it->bmpIcon)
			it->bmpIcon.DeleteObject();
	}

	s_vecFavoritesItem.clear();
}

void	CFavoritesOption::LoadConfig()
{
	CString FavoritesConfigPath = Misc::GetExeDirectory() + _T("Favorites.xml");

	CleanFavoritesItem();
	s_vecFavoritesItem.reserve(20);

	wptree	pt;
	std::wifstream 	configstream(FavoritesConfigPath, std::ios::in);
	if (!configstream)
		return ;
	configstream.imbue(std::locale(std::locale(), new std::codecvt_utf8_utf16<wchar_t>));
	read_xml(configstream, pt);

	if (auto ptChild = pt.get_child_optional(L"Favorites")) {
		for (auto it = ptChild->begin(); it != ptChild->end(); ++it) {
			wptree ptItem = it->second;
			FavoritesItem	itemdata;
			itemdata.strTitle = ptItem.get<std::wstring>(L"title", L"").c_str();
			itemdata.strPath	= ptItem.get<std::wstring>(L"path", L"").c_str();
			if (itemdata.strPath != FAVORITESSEPSTRING) {
				CStringA base64data = ptItem.get<std::wstring>(L"ITEMIDLIST", L"").c_str();
				if (base64data.GetLength() > 0) {
					int nLength = ::Base64DecodeGetRequiredLength(base64data.GetLength());
					itemdata.pidl	= (LPITEMIDLIST)::CoTaskMemAlloc(nLength);
					::SecureZeroMemory((PVOID)itemdata.pidl, nLength);
					::Base64Decode(base64data, base64data.GetLength(), (BYTE*)itemdata.pidl, &nLength);
				}
				CIcon icon = ShellWrap::CreateIconFromIDList(itemdata.pidl);
				if (icon.m_hIcon) {
					itemdata.bmpIcon	= Misc::CreateBitmapFromHICON(icon);
				}
			}
			s_vecFavoritesItem.push_back(itemdata);
		}
	}

}

void	CFavoritesOption::SaveConfig()
{
	wptree	pt;
	wptree	ptChild;
	for (auto it = s_vecFavoritesItem.begin(); it != s_vecFavoritesItem.end(); ++it) {
		wptree ptItem;
		ptItem.add(L"title", (LPCTSTR)it->strTitle);
		ptItem.add(L"path"	, (LPCTSTR)it->strPath);
		UINT uSize = ::ILGetSize(it->pidl);
		CStringA	strbuff;
		int			buffSize = ::Base64EncodeGetRequiredLength((int)uSize, ATL_BASE64_FLAG_NOCRLF | ATL_BASE64_FLAG_NOPAD);
		if (::Base64Encode((BYTE*)it->pidl, uSize, strbuff.GetBuffer(buffSize), &buffSize, ATL_BASE64_FLAG_NOCRLF | ATL_BASE64_FLAG_NOPAD)) {
			CString strWide = strbuff;
			strWide.SetAt(buffSize - 1, L'\0');
			ptItem.add(L"ITEMIDLIST", (LPCTSTR)strWide);
		}
		ptChild.add_child(L"item", ptItem);
	}
	pt.add_child(L"Favorites", ptChild);

	CString FavoritesConfigPath = Misc::GetExeDirectory() + _T("Favorites.xml");
	std::wofstream 	ostream(FavoritesConfigPath, std::ios::out | std::ios::trunc);
	if (!ostream) {
		MessageBox(NULL, _T("Favorites.xmlのオープンに失敗"), NULL, MB_ICONERROR);
		return ;
	}

	ostream.imbue(std::locale(std::locale(), new std::codecvt_utf8_utf16<wchar_t>));
	write_xml(ostream, pt, xml_writer_make_settings(L' ', 2, xml_parser::widen<wchar_t>("utf-8")));

}

//////////////////////////////////////////////
/// お気に入り
class CFavoritesPropertyPage : 
	public CPropertyPageImpl<CFavoritesPropertyPage>,
	public CWinDataExchange<CFavoritesPropertyPage>,
	protected CFavoritesOption
{
public:
	enum { IDD = IDD_FAVORITES };

	// DDX map
    BEGIN_DDX_MAP(CThumbnailTooltipPropertyPage)
		DDX_CONTROL_HANDLE(IDC_LIST_FAVORITS, m_ListFavorites )
		DDX_TEXT(IDC_EDIT_TITLE	, m_strTitle)
		DDX_TEXT(IDC_EDIT_PATH	, m_strPath)
    END_DDX_MAP()
	

	// Message map
	BEGIN_MSG_MAP_EX( CFavoritesPropertyPage )
		MSG_WM_INITDIALOG( OnInitDialog )
		COMMAND_HANDLER_EX(IDC_LIST_FAVORITS, LBN_SELCHANGE, OnListSelChange )
		COMMAND_ID_HANDLER_EX( IDC_BUTTON_UP, OnUp )
		COMMAND_ID_HANDLER_EX( IDC_BUTTON_DOWN, OnDown )
		COMMAND_ID_HANDLER_EX( IDC_BUTTON_ADDSEP, OnAddSep )
		COMMAND_ID_HANDLER_EX( IDC_BUTTON_DEL, OnDel )
		COMMAND_ID_HANDLER_EX( IDC_BUTTON_ADD, OnAdd )
		COMMAND_ID_HANDLER_EX( IDC_BUTTON_APPLY, OnButtonApply )
		CHAIN_MSG_MAP( CPropertyPageImpl<CFavoritesPropertyPage> )
	END_MSG_MAP()

	BOOL OnInitDialog(CWindow wndFocus, LPARAM lInitParam)
	{
		DoDataExchange(DDX_LOAD);
		for (auto it = s_vecFavoritesItem.begin(); it != s_vecFavoritesItem.end(); ++it) {
			m_ListFavorites.AddString(it->strTitle);
		}
		return 0;
	}

	BOOL OnApply()
	{
		if (!DoDataExchange(DDX_SAVE))
			return FALSE;

		SaveConfig();	/* 保存 */

		return TRUE;
	}

	void OnListSelChange(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1)
			return ;
		m_strTitle	= s_vecFavoritesItem[nIndex].strTitle;
		m_strPath	= s_vecFavoritesItem[nIndex].strPath;
		DoDataExchange(DDX_LOAD);
	}

	void OnUp(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1 || nIndex == 0)
			return ;
		auto item = s_vecFavoritesItem[nIndex];
		s_vecFavoritesItem.erase(s_vecFavoritesItem.begin() + nIndex);
		s_vecFavoritesItem.insert(s_vecFavoritesItem.begin() + nIndex - 1, item);

		m_ListFavorites.DeleteString(nIndex);
		m_ListFavorites.InsertString(nIndex - 1, item.strTitle);
		m_ListFavorites.SetCurSel(nIndex - 1);
	}

	void OnDown(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1 || nIndex == m_ListFavorites.GetCount() - 1)
			return ;
		auto item = s_vecFavoritesItem[nIndex];
		s_vecFavoritesItem.insert(s_vecFavoritesItem.begin() + nIndex + 2, item);
		s_vecFavoritesItem.erase(s_vecFavoritesItem.begin() + nIndex);

		m_ListFavorites.InsertString(nIndex + 2, item.strTitle);
		m_ListFavorites.DeleteString(nIndex);
		m_ListFavorites.SetCurSel(nIndex + 1);
	}

	void OnAddSep(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1)
			return ;
		FavoritesItem	item;
		item.strTitle	= FAVORITESSEPSTRING;
		item.strPath	= FAVORITESSEPSTRING;
		s_vecFavoritesItem.insert(s_vecFavoritesItem.begin() + nIndex, item);

		m_ListFavorites.InsertString(nIndex, FAVORITESSEPSTRING);
	}

	void OnDel(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1)
			return ;
		if (s_vecFavoritesItem[nIndex].bmpIcon) 
			s_vecFavoritesItem[nIndex].bmpIcon.DeleteObject();
		if (s_vecFavoritesItem[nIndex].pidl)
			::ILFree(s_vecFavoritesItem[nIndex].pidl);
		s_vecFavoritesItem.erase(s_vecFavoritesItem.begin() + nIndex);

		m_ListFavorites.DeleteString(nIndex);
		if (nIndex == m_ListFavorites.GetCount())
			--nIndex;
		m_ListFavorites.SetCurSel(nIndex);
		OnListSelChange(0, 0, NULL);
		if (nIndex == -1) {
			m_strTitle	= _T("");
			m_strPath	= _T("");
			DoDataExchange(DDX_LOAD);
		}
	}

	void OnAdd(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1)
			nIndex = 0;
		DoDataExchange(DDX_SAVE);
		if (m_strTitle.IsEmpty() || m_strPath.IsEmpty()) {
			MessageBox(_T("名前もしくはパスが設定されていません"), NULL, MB_ICONWARNING);
			m_strTitle = _T("");
			m_strPath = _T("");
			DoDataExchange(DDX_LOAD);
			return ;
		}
		if (m_strPath == FAVORITESSEPSTRING) {
			OnAddSep(0, 0, NULL);
			return ;
		}
		FavoritesItem	item;
		item.strTitle	= m_strTitle;
		item.strPath	= m_strPath;
		item.pidl		= ShellWrap::CreateIDListFromFullPath(m_strPath);
		if (item.pidl) {
			CIcon icon = ShellWrap::CreateIconFromIDList(item.pidl);
			if (icon.m_hIcon)
				item.bmpIcon	= Misc::CreateBitmapFromHICON(icon);
		} else {
			MessageBox(item.strPath + _T(" は存在しないパスです"), NULL, MB_ICONWARNING);
		}
		s_vecFavoritesItem.insert(s_vecFavoritesItem.begin() + nIndex + 1, item);

		int nInsert = m_ListFavorites.InsertString(nIndex + 1, m_strTitle);
		m_ListFavorites.SetCurSel(nInsert);
	}

	void OnButtonApply(UINT uNotifyCode, int nID, CWindow wndCtl)
	{
		int nIndex = m_ListFavorites.GetCurSel();
		if (nIndex == -1)
			return ;

		DoDataExchange(DDX_SAVE);
		if (m_strTitle.IsEmpty() || m_strPath.IsEmpty()) {
			MessageBox(_T("名前もしくはパスが設定されていません"), NULL, MB_ICONWARNING);
			m_strTitle = _T("");
			m_strPath = _T("");
			DoDataExchange(DDX_LOAD);
			return ;
		}

		OnAdd(0, 0, NULL);
		m_ListFavorites.SetCurSel(nIndex);
		OnDel(0, 0, NULL);		
	}

private:
	CListBox	m_ListFavorites;
	CString	m_strTitle;
	CString m_strPath;

};



///////////////////////////////////////////////////////////
///  オプションダイアログ


/// プロパティシートを表示
INT_PTR	CExpTabBarOption::Show(HWND hWndParent)
{
	SetTitle(_T("ExpTabBarの設定"));
	m_psh.dwFlags |= PSH_NOAPPLYNOW | PSH_NOCONTEXTHELP;

	CTabBarOptionPropertyPage	tabPage;
	AddPage(tabPage);
	CThumbnailTooltipPropertyPage	thumbtipPage;
	AddPage(thumbtipPage);
	CFavoritesPropertyPage	favoritesPage;
	AddPage(favoritesPage);
	
	return DoModal(hWndParent);
}


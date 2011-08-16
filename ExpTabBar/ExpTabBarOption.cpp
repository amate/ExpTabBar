/**
*	@file	ExpBabBarOption.cpp
*	@brief	オプション
*/

#include "stdafx.h"
#include "ExpTabBarOption.h"
#include "resource.h"
#include <atlddx.h>
#include <boost/property_tree/ptree.hpp>
#include <boost/property_tree/ini_parser.hpp>
#include <boost/optional.hpp>
#include "DonutFunc.h"

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
}

/// 設定を保存する
void	CTabBarConfig::SaveConfig()
{
	ptree pt;
	std::ofstream	inistream(g_szIniFileName);

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

	write_ini(inistream, pt);
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







///////////////////////////////////////////////////////////
///  オプションダイアログ


/// プロパティシートを表示
INT_PTR	CExpTabBarOption::Show(HWND hWndParent)
{
	SetTitle(_T("ExpTabBarの設定"));
	m_psh.dwFlags |= PSH_NOAPPLYNOW | PSH_NOCONTEXTHELP;

	CTabBarOptionPropertyPage	tabPage;
	AddPage(tabPage);
	
	return DoModal(hWndParent);
}

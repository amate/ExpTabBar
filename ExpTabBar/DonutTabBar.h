// DonutTabBar.h
#pragma once

#include "OleDragDropTabCtrl.h"
#include "HlinkDataObject.h"

#define	WM_NEWTABBUTTON	(WM_APP + 1)


// Forward Declare
class CDonutTabBar;

////////////////////////////////////////////////////////////
// CNotifyWindow	: 通知用のダミーのウィンドウ

class CNotifyWindow : public CFrameWindowImpl<CNotifyWindow>
{
public:
	DECLARE_FRAME_WND_CLASS(_T("ExpTabBar_NotifyWindow"), NULL)

	// Constructor
	CNotifyWindow(CDonutTabBar* p);

	BEGIN_MSG_MAP_EX( CNotifyWindow )
		MSG_WM_COPYDATA( OnCopyData )
		//CHAIN_MSG_MAP( CFrameWindowImpl<CNotifyWindow> )
	END_MSG_MAP()

	BOOL OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct);

private:
	CDonutTabBar*	m_pTabBar;
};





// Extended styles
enum EMtb_Ex {
	MTB_EX_ADDLEFT				= 0x00000001L,	// (一番)左に追加
	MTB_EX_RIGHTCLICKCLOSE		= 0x00000002L,	// 右クリックで閉じる
	MTB_EX_DELAYED				= 0x00000004L,
	MTB_EX_MULTILINE			= 0x00000008L,	// 複数列表示
	MTB_EX_DOUBLECLICKCLOSE 	= 0x00000010L,	// ダブルクリックで閉じる
	MTB_EX_XCLICKCLOSE			= 0x00000020L,	// ホイールクリックで閉じる
	MTB_EX_RIGHTCLICKREFRESH	= 0x00000040L,	// 右クリックで更新
	MTB_EX_DOUBLECLICKREFRESH	= 0x00000080L,	// ダブルクリックで更新
	MTB_EX_XCLICKREFRESH		= 0x00000100L,	// ホイールクリックで更新
	MTB_EX_RIGHTACTIVEONCLOSE	= 0x00000200L,	// 閉じたとき右側をアクティブに
	MTB_EX_LEFTACTIVEONCLOSE	= 0x00000400L,	// 閉じたとき左側をアクティブに
	MTB_EX_ADDRIGHTACTIVE		= 0x00000800L,	// アクティブの右側に追加
	MTB_EX_ADDLEFTACTIVE		= 0x00001000L,	// アクティブの左側に追加
	MTB_EX_WHEEL				= 0x00002000L,	// ホイールでタブ切り替え？
	MTB_EX_FIXEDSIZE			= 0x00004000L,	// タブの大きさを固定する
	MTB_EX_ANCHORCOLOR			= 0x00008000L,	// 表示/未表示を色分けする
	// UDT DGSTR
	MTB_EX_DOUBLECLICKNLOCK 	= 0x00010000L,	// ダブルクリックでナビゲートロック
	MTB_EX_XCLICKNLOCK			= 0x00020000L,	// ホイールクリックでナビゲートロック
	MTB_EX_ADDLINKACTIVERIGHT	= 0x00040000L,	// リンクはアクティブなタブの右側に追加
	// minit
	MTB_EX_DOUBLECLICKCOMMAND	= 0x00080000L,
	MTB_EX_XCLICKCOMMAND		= 0x00100000L,
	MTB_EX_RIGHTCLICKCOMMAND	= 0x00200000L,
	MTB_EX_MOUSEDOWNSELECT		= 0x00400000L,	// クリックでタブ選択(Drag&Drop不可)
	// +mod
	MTB_EX_CTRLTAB_MDI			= 0x00800000L,
	
	// 初期設定
	MTB_EX_DEFAULT_BITS	=	
					MTB_EX_MULTILINE			//   複数列表示
				  | MTB_EX_DOUBLECLICKCLOSE		// | ダブルクリックで閉じる
				  | MTB_EX_XCLICKCLOSE			// | ホイールクリックで閉じる
				  | MTB_EX_LEFTACTIVEONCLOSE	// | 閉じたとき左側をアクティブに
				  | MTB_EX_ADDLINKACTIVERIGHT	// | リンクはアクティブなタブの右側に追加
};



// A_ONはGetLinkStateで消されるB_ONは残る
enum ELinkState {
	LINKSTATE_OFF		= 0,
	LINKSTATE_A_ON		= 1,
	LINKSTATE_B_ON		= 2,
};


///////////////////////////////////////////////////////////////////////////////////////
// CDonutTabBar

class CDonutTabBar : public COleDragDropTabCtrl<CDonutTabBar>
{
public:
	DECLARE_WND_CLASS_EX(_T("DonutTabBarCtrl"), CS_DBLCLKS, COLOR_BTNFACE)

	// コンストラクタ
	CDonutTabBar();
	~CDonutTabBar();

	void	Initialize(IUnknown* punk);

	void	SaveAllTab();
	void	RestoreAllTab();

	void	SaveHistory();
	void	RestoreHistory();

	int		OnTabCreate(LPITEMIDLIST pidl, bool bAddLast = false, bool bInsert = false);
	void	OnTabDestroy(int nDestroyIndex);

	int		AddTabItem(CTabItem& item);

	void	NavigateLockTab(int nIndex, bool bOn);


	DWORD	GetTabExtendedStyle() const { return m_dwExtendedStyle; }

	void	ExternalOpen(LPCTSTR strFullPath);

	// overrides
	void	OnSetCurSel(int nIndex, int nOldIndex);
	HRESULT OnGetTabCtrlDataObject(CSimpleArray<int> arrIndex, IDataObject **ppDataObject);
	void	OnVoidTabRemove(const CSimpleArray<int>& arrCurDragItems);

	DROPEFFECT	OnDragEnter(IDataObject *pDataObject, DWORD dwKeyState, CPoint point);
	DROPEFFECT	OnDragOver (IDataObject *pDataObject, DWORD dwKeyState, CPoint point, DROPEFFECT dropOkEffect);
	DROPEFFECT	OnDrop	   (IDataObject *pDataObject, DROPEFFECT dropEffect, DROPEFFECT dropEffectList, CPoint point);

	// notify
	void	RefreshTab(LPCTSTR title);
	void	NavigateComplete2(LPCTSTR strURL);
	void	DocumentComplete();

	// メッセージマップ
	BEGIN_MSG_MAP_EX( CDonutTabBar )
		MSG_WM_CREATE		( OnCreate )
		MSG_WM_DESTROY		( OnDestroy )
		MSG_WM_RBUTTONUP	( OnRButtonUp	 )		// 右クリック
		MSG_WM_INITMENUPOPUP( OnInitMenuPopup )
		MSG_WM_MENUSELECT	( OnMenuSelect )
		MSG_WM_LBUTTONDBLCLK( OnLButtonDblClk)		// ダブルクリック
		MSG_WM_MBUTTONUP	( OnMButtonUp	 )		// ホイールクリック
		MESSAGE_HANDLER_EX( WM_NEWTABBUTTON, OnNewTabButton )
		COMMAND_ID_HANDLER_EX( ID_TABCLOSE			, OnTabClose )
		COMMAND_ID_HANDLER_EX( ID_RIGHTALLCLOSE		, OnRightAllClose )
		COMMAND_ID_HANDLER_EX( ID_LEFTALLCLOSE		, OnLeftAllClose )
		COMMAND_ID_HANDLER_EX( ID_EXCEPTCURTABCLOSE	, OnExceptCurTabClose )
		COMMAND_ID_HANDLER_EX( ID_OPEN_UPFOLDER		, OnOpenUpFolder )
		COMMAND_ID_HANDLER_EX( ID_NAVIGATELOCK		, OnNavigateLock )
		COMMAND_RANGE_HANDLER_EX( ID_RECENTCLOSED_FIRST, ID_RECENTCLOSED_LAST, OnClosedTabCreate )
		NOTIFY_CODE_HANDLER( TTN_GETDISPINFO, OnTooltipGetDispInfo )
		CHAIN_MSG_MAP( COleDragDropTabCtrl<CDonutTabBar> )
	END_MSG_MAP()


	int		OnCreate(LPCREATESTRUCT lpCreateStruct);
	void	OnDestroy();
	void	OnRButtonUp(UINT nFlags, CPoint point);
	void	OnInitMenuPopup(CMenuHandle menuPopup, UINT nIndex, BOOL bSysMenu);
	void	OnMenuSelect(UINT nItemID, UINT nFlags, CMenuHandle menu);
	void	OnLButtonDblClk(UINT nFlags, CPoint point);
	void	OnMButtonUp(UINT nFlags, CPoint point);
	LRESULT OnNewTabButton(UINT uMsg, WPARAM wParam, LPARAM lParam);

	// command
	void	OnTabClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnRightAllClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnLeftAllClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnExceptCurTabClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnOpenUpFolder(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnNavigateLock(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnClosedTabCreate(UINT uNotifyCode, int nID, CWindow wndCtl);

	LRESULT OnTooltipGetDispInfo(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);


private:
	int		_ManageInsert();
	int		_ManageClose(int nActiveIndex);

	bool	_SaveTravelLog(int nIndex);
	bool	_RestoreTravelLog(int nIndex);
	bool	_RemoveCurAllTravelLog();
	void	_RemoveOneBackLog();
	CString	_OneBackLog();

	int		_IDListIsEqualIndex(LPITEMIDLIST pidl);

	static void _thread_ChangeSelectedItem(LPVOID p);
	void	_ChangeSelectedItem();

	static void _thread_VoidTabRemove(LPVOID p);
	void	_VoidTabRemove();

	void	_AddHistory(int nDestroy);
	void	_DeleteHistory(int nIndex);

	void	_ClearSearchText();
	
	void	_SaveSelectedIndex(int nIndex);

	// Data members
	DWORD		m_dwExtendedStyle;
	CMenu		m_menuPopup;
	int			m_ClickedIndex;				// 右クリックメニューを表示したタブ

	CComPtr<IShellBrowser>	m_spShellBrowser;
	CComPtr<ITravelLogStg>	m_spTravelLogStg;
	CComPtr<ISearchBoxInfo>	m_spSearchBoxInfo;

	bool		m_bLeftButton;
	CString		m_strDraggingDrive;
	bool		m_bDragAccept;
	bool		m_bTabChanging;			// タブ切り替え中かどうか
	bool		m_bTabChanged;
	bool		m_bNavigateLockOpening;
	bool		m_bSaveAllTab;			// タブを保存するかどうか

	CSimpleArray<int>	m_arrDragItems;

	struct HistoryItem {
		CString			strTitle;
		CString			strFullPath;
		LPITEMIDLIST	pidl;
		TRAVELLOG*		pLogBack;
		TRAVELLOG*		pLogFore;
		HBITMAP			hbmp;

		HistoryItem()
			: pidl(NULL)
			, pLogBack(NULL)
			, pLogFore(NULL)
			, hbmp(NULL) { }
	};

	vector<HistoryItem>	m_vecHistoryItem;
	CMenu			m_menuHistory;
	CToolTipCtrl	m_tipHistroy;
	int				m_nMaxHistory;

	HWND			m_hSearch;

	CNotifyWindow	m_wndNotify;
	int				m_nInsertIndex;
};













































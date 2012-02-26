// DonutTabBar.h
#pragma once

#include "OleDragDropTabCtrl.h"
#include "HlinkDataObject.h"

//#define	WM_NEWTABBUTTON	(WM_APP + 1)


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

	// Constants
	enum { AutoSaveTimerID = 1, AutoSaveInterval =  60 * 1000 };

	// コンストラクタ
	CDonutTabBar();
	~CDonutTabBar();

	void	Initialize(IUnknown* punk);

	void	SaveAllTab();
	void	RestoreAllTab();

	void	SaveHistory();
	void	RestoreHistory();

	int		OnTabCreate(LPITEMIDLIST pidl, bool bAddLast = false, bool bInsert = false, bool bLink = false);
	void	OnTabDestroy(int nDestroyIndex);

	int		AddTabItem(CTabItem& item);

	void	NavigateLockTab(int nIndex, bool bOn);

	void	ExternalOpen(LPITEMIDLIST pidl);
	void	ExternalOpen(LPCTSTR strFullPath);

	// overrides
	void	OnSetCurSel(int nIndex, int nOldIndex);
	HRESULT OnGetTabCtrlDataObject(CSimpleArray<int> arrIndex, IDataObject **ppDataObject);
	void	OnVoidTabRemove(const CSimpleArray<int>& arrCurDragItems);

	DROPEFFECT	OnDragEnter(IDataObject *pDataObject, DWORD dwKeyState, CPoint point);
	DROPEFFECT	OnDragOver (IDataObject *pDataObject, DWORD dwKeyState, CPoint point, DROPEFFECT dropOkEffect);
	DROPEFFECT	OnDrop	   (IDataObject *pDataObject, DROPEFFECT dropEffect, DROPEFFECT dropEffectList, CPoint point);
	void		OnDragLeave();

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
		MSG_WM_TIMER		( OnTimer )
		MSG_WM_MOUSEWHEEL	( OnMouseWheel )
		COMMAND_ID_HANDLER_EX( ID_TABCLOSE			, OnTabClose )
		COMMAND_ID_HANDLER_EX( ID_RIGHTALLCLOSE		, OnRightAllClose )
		COMMAND_ID_HANDLER_EX( ID_LEFTALLCLOSE		, OnLeftAllClose )
		COMMAND_ID_HANDLER_EX( ID_EXCEPTCURTABCLOSE	, OnExceptCurTabClose )
		COMMAND_ID_HANDLER_EX( ID_OPEN_UPFOLDER		, OnOpenUpFolder )
		COMMAND_ID_HANDLER_EX( ID_ADD_FAVORITES		, OnAddFavorites )
		COMMAND_ID_HANDLER_EX( ID_NAVIGATELOCK		, OnNavigateLock )
		COMMAND_ID_HANDLER_EX( ID_OPTION			, OnOpenOption )
		COMMAND_RANGE_HANDLER_EX( ID_RECENTCLOSED_FIRST, ID_RECENTCLOSED_LAST, OnClosedTabCreate )
		COMMAND_RANGE_HANDLER_EX( ID_FAVORITES_FIRST	, ID_FAVORITES_LAST, OnFavoritesOpen )
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
	void	OnTimer(UINT_PTR nIDEvent);
	BOOL	OnMouseWheel(UINT nFlags, short zDelta, CPoint pt);

	// command
	void	OnTabClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnRightAllClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnLeftAllClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnExceptCurTabClose(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnOpenUpFolder(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnAddFavorites(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnNavigateLock(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnOpenOption(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnClosedTabCreate(UINT uNotifyCode, int nID, CWindow wndCtl);
	void	OnFavoritesOpen(UINT uNotifyCode, int nID, CWindow wndCtl);

	LRESULT OnTooltipGetDispInfo(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);


private:
	int		_ManageInsert(bool bLink);
	int		_ManageClose(int nActiveIndex);

	bool	_SaveTravelLog(int nIndex);
	bool	_RestoreTravelLog(int nIndex);
	bool	_RemoveCurAllTravelLog();
	void	_RemoveOneBackLog();
	CString	_OneBackLog();

	int		_IDListIsEqualIndex(LPITEMIDLIST pidl);

	void	_threadVoidTabRemove();

	void	_AddHistory(int nDestroy);
	void	_DeleteHistory(int nIndex);

	void	_ClearSearchText();
	
	void	_SaveSelectedIndex(int nIndex);

	void	_threadPerformSHFileOperation(LPITEMIDLIST pidlTo, IDataObject*	pDataObject, bool bMove);

	void	_ExeCommand(int nIndex, int Command);

	// Data members
	CMenu		m_menuPopup;
	int			m_ClickedIndex;				// 右クリックメニューを表示したタブ

	CComPtr<IShellBrowser>	m_spShellBrowser;
	CComPtr<ITravelLogStg>	m_spTravelLogStg;
	CComPtr<ISearchBoxInfo>	m_spSearchBoxInfo;

	// for DragDrop
	bool		m_bLeftButton;
	int			m_DragItemDriveNumber;
	bool		m_bDragAccept;
	bool		m_bDragItemIncludeFolder;
	bool		m_bDragItemIsRoot;


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
	CMenu			m_menuFavorites;
	CToolTipCtrl	m_tipHistroy;

	HWND			m_hSearch;

	CNotifyWindow	m_wndNotify;
	int				m_nInsertIndex;

	CToolTipCtrl	m_tipDragOver;
};













































// TabCtrlImpl.h
#pragma once

#include "TabCtrl.h"

//////////////////////////////////////////////////////
// CTabCtrlImpl

class CTabCtrlImpl 
	: public CDoubleBufferWindowImpl<CTabCtrlImpl>
	, public CThemeImpl<CTabCtrlImpl>
	, public CTrackMouseLeave<CTabCtrlImpl>
{
	friend class CTabSkin;
public:

	void	ReloadSkinData();

	CTabItem& GetItem(int nIndex) { return m_items.at(nIndex); }

	// overridables
	virtual void OnSetCurSel(int nIndex, int nOldIndex) = 0;

	// overrides
	void	OnTrackMouseMove(UINT nFlags, CPoint pt);
	void	OnTrackMouseLeave();

	void	DoPaint(CDCHandle dc);

	// attributes
	int		GetItemCount() const { return (int) m_items.size(); }

	DWORD	GetTabStyle() const { return m_TabStyle; }
	void	SetTabStyle(DWORD style) { m_TabStyle = style; }

	void	SetItemSize(const CSize& size);

	int		GetItemHeight() const;

	void	GetCurMultiSel(CSimpleArray<int> &arrDest, bool bIncludeCurSel = true);
	void	GetCurMultiSelEx(CSimpleArray<int>& arrDest, int nIndex);
	bool	CanScrollItem(bool bRight) const;

	int		ReplaceIcon(int nIndex, HICON hIcon);
	int		AddIcon(HICON hIcon);
	HICON	GetIcon(int nImgIndex) { return m_imgs.GetIcon(nImgIndex); }

	// oparation
	int		SetCurSel(int nIndex, bool bClicked = false);
	int		GetCurSel();

	int		HitTest(const CPoint& point);

	int		InsertItem(int nIndex, const CTabItem &item);
	int		AddItem(const CTabItem &item);

	void	DeleteItem(int nIndex, bool bMoveNow = false);
	void	DeleteItems(const CSimpleArray<int> &arrSrcs);

	bool	MoveItems(int nDestIndex, const CSimpleArray<int> &arrSrcs);

	bool	ScrollItem(bool bRight = true);

	// CTabItem
	WORD	GetItemState(int nIndex) const { return m_items.at(nIndex).m_fsState; }

	LPCTSTR	GetItemText(int nIndex) const { return m_items.at(nIndex).m_strItem; }
	void	SetItemText(int nIndex, LPCTSTR strText);

	int		GetItemImageIndex(int nIndex) const { return m_items.at(nIndex).m_nImgIndex; }
	void	SetItemImageIndex(int nIndex, int nImgIndex);

	LPITEMIDLIST GetItemIDList(int nIndex) const { return m_items.at(nIndex).m_pidl; }
	void	SetItemIDList(int nIndex, LPITEMIDLIST pidl);

	CString	GetItemFullPath(int nIndex) const { return m_items.at(nIndex).m_strFullPath; }
	void	SetItemFullPath(int nIndex, LPCTSTR FullPath) { m_items.at(nIndex).m_strFullPath = FullPath; }

	int		GetItemSelectedIndex(int nIndex) const { return m_items.at(nIndex).m_nSelectedIndex; }
	void	SetItemSelectedIndex(int nIndex, int nSelect) { m_items.at(nIndex).m_nSelectedIndex = nSelect; }


	// メッセージマップ
	BEGIN_MSG_MAP_EX(CTabCtrlImpl)
		CHAIN_MSG_MAP( CThemeImpl<CTabCtrlImpl> )
		MSG_WM_CREATE			( OnCreate )
		MSG_WM_DESTROY			( OnDestroy )
		MSG_WM_WINDOWPOSCHANGING( OnWindowPosChanging )
		MSG_WM_LBUTTONDOWN		( OnLButtonDown )
		MSG_WM_LBUTTONUP		( OnLButtonUp )
//		COMMAND_ID_HANDLER_EX( COMMAND_ID_DROPBUTTON, OnPushDropButton	)
		NOTIFY_CODE_HANDLER( UDN_DELTAPOS		, OnDeltaPos	)
		NOTIFY_CODE_HANDLER( TTN_GETDISPINFO	, OnGetDispInfo	)
		CHAIN_MSG_MAP( CTrackMouseLeave<CTabCtrlImpl> )
		CHAIN_MSG_MAP( CDoubleBufferWindowImpl<CTabCtrlImpl> )
//		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()

	int 	OnCreate(LPCREATESTRUCT lpCreateStruct);
	void	OnDestroy();
	void	OnWindowPosChanging(LPWINDOWPOS lpWndPos);
	void	OnLButtonDown(UINT nFlags, CPoint point);
	void	OnLButtonUp(UINT nFlags, CPoint point);
	LRESULT OnDeltaPos(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);
	LRESULT OnGetDispInfo(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

protected:
	// コンストラクタ
	CTabCtrlImpl();
	virtual ~CTabCtrlImpl();

	void	_InitToolTip();

	bool	_IsValidIndex(int nIndex) const;

	int		_GetRequiredHeight();
	CRect	_MeasureItemRect(const CString &strText);

	void	_HotItem(int nNewHotIndex = -1);
	void	_PressItem(int nNewPressedIndex = -1);

	bool	_FindIndexFromCurMultiSelected(int nIndex);
	void	_ResetMultiSelectedItems();

	void	_ShowOrHideUpDownCtrl(const CRect &rcClient);
	void	_ScrollOpposite(int nNewSel, bool bClicked);

	void	_RefreshBandInfo();
	void	_UpdateSingleLineLayout(int nFirstIndex, bool bForce = false);
	void	_UpdateMultiLineLayout(int nWidth);
	void	_UpdateLayout();

	const CRect	_InflateGapWidth(const CRect &rc) const;

protected:

	enum _TabCtrlDrawConstants
	{
		s_kcxTextMargin = 7,
		s_kcyTextMargin = 3,	// タブとテキスト間の上下の余白
		s_kcxGap		= 2,
		s_kcyGap		= -2,	// タブの上下間の幅
		s_kSideMargin	= 2,	// 左右の余白
		s_kcxSeparator	= 2,
		s_kcyTabIcon	= 4,	// タブとアイコンの上下の余白

		s_kcxUpDown 	= 28,
		s_kcyUpDown 	= 14,
	};

	vector<CTabItem>	m_items;
	CImageList			m_imgs;

	CUpDownCtrl			m_wndUpDown;
	CDropDownButtonCtrl	m_wndDropBtn;

	int					m_nFirstIndexOnSingleLine;	// for single line

private:
	bool					m_bLockRefreshBandInfo;

	DWORD					m_TabStyle;

	int						m_nHotIndex;				// Current hot index
	int						m_nPressedIndex;			// Current capturing index
	int						m_nActiveIndex;				// Current Active index

	CSimpleArray<CPoint>	m_arrSeparators;			// Point of separator

	CToolTipCtrl			m_tip;
	CString					m_strTooltip;

	CTabSkin*				m_pTabSkin;
	
};
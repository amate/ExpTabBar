// TabCtrl.h

#pragma once

#include "DropDownButton.h"
#include "DonutFunc.h"

#undef max
#undef min

typedef std::vector<std::pair<CString, CString> >	TRAVELLOG;

namespace MTL {

// for debug
#ifdef _DEBUG
const bool _Mtl_TabCtrl2_traceOn = false;
	# define TCTRACE	if (_Mtl_TabCtrl2_traceOn)	ATLTRACE
#else
	# define TCTRACE
#endif

using std::vector;

#define	CXICON	16
#define CYICON	16


/////////////////////////////////////////////////////////////////////////////
// Tab Bars

// Extended styles
enum ETab2_Ex {
	TAB2_EX_TRANSPARENT 	= 0x00000001L,	// not supported yet
	TAB2_EX_SHAREIMGLIST	= 0x00000002L,
	TAB2_EX_MULTILINE		= 0x00000004L,	// 複数列表示
	TAB2_EX_FIXEDSIZE		= 0x00000008L,	// タブの大きさを固定する
	TAB2_EX_SUPPORTREBAR	= 0x00000010L,
	TAB2_EX_SCROLLOPPOSITE	= 0x00000020L,
	TAB2_EX_ANCHORCOLOR 	= 0x00000040L,	// タブを表示、未表示で色を変える
	TAB2_EX_MOUSEDOWNSELECT = 0x00000080L,	// タブを左クリックしたときにタブを切り替える
											// ※Drag&Dropでタブを並び変えられない

	TAB2_EX_DEFAULT_BITS	= 
					TAB2_EX_TRANSPARENT 
				  | TAB2_EX_MULTILINE		// 複数列表示
				  | TAB2_EX_FIXEDSIZE		// タブの大きさを固定する
				  | TAB2_EX_SUPPORTREBAR 
				  | TAB2_EX_SCROLLOPPOSITE
};

// SkinOption.hから
enum ESkn_Tab_Style {
	SKN_TAB_STYLE_DEFAULT		= 0,
	SKN_TAB_STYLE_CLASSIC		= 1,
	SKN_TAB_STYLE_THEME 		= 2,
};


/////////////////////////////////////////////////////////////////////////////
// CTabButton

// CTabButton state flags
enum ETcistate {
	TCISTATE_HIDDEN 			= 0x001,
	// standard text state
	TCISTATE_ENABLED			= 0x002,
	TCISTATE_INACTIVE			= 0x004,	// タブが未選択
	// select or not
	TCISTATE_SELECTED			= 0x008, 	// ordianry selected
	TCISTATE_MSELECTED			= 0x010, 	// 複数選択されている
	// event state
	TCISTATE_PRESSED			= 0x020, 	// mouse pressed
	TCISTATE_HOT				= 0x040, 	// mouse above item
	TCISTATE_MOUSEAWAYCAPTURED	= 0x080,	 // mouse away but captured

	TCISTATE_NAVIGATELOCK		= 0x100,
};

// Command ID
enum { COMMAND_ID_DROPBUTTON = 0x00000001L, };



/////////////////////////////////////////////////////////////////////////////
// Forward declarations

//class	CTabCtrl2;

class	CTabItem;


/////////////////////////////////////////////////////////////////////////////
// CTabSkin

class CTabSkin
{
	friend class CTabItem;

	enum { s_kcxIconGap = 5 };

private:
	CFont		m_font;

	COLORREF	m_colText;
	COLORREF	m_colActive;
	COLORREF	m_colInActive;
	COLORREF	m_colDisable;
	COLORREF	m_colDisableHi;
	COLORREF	m_colBtnFace;
	COLORREF	m_colBtnHi;

	CImageList	m_imgLock;

protected:
	// コンストラクタ
	CTabSkin() { _LoadTabTextSetting(); _LoadImage(); }
public:
	virtual ~CTabSkin() { m_imgLock.Destroy(); }

	void	Update(CDCHandle dc, HIMAGELIST hImgList, const CTabItem& item, bool bAnchorColor);

	int		GetFontHeight() const;
	HFONT	GetFontHandle() { return m_font.m_hFont; }

private:
	void	_GetSystemFontSetting(LOGFONT* plf);
	void	_LoadTabTextSetting();
	void	_LoadImage();

	void	_DrawText(CDCHandle dc, CPoint ptOffset, const CTabItem& item, bool bAnchorColor);

	virtual void  _DrawSkinCur(CDCHandle dc, CRect rcItem) = 0;	// Active
	virtual void  _DrawSkinNone(CDCHandle dc, CRect rcItem) = 0;	// Normal
	virtual void  _DrawSkinSel(CDCHandle dc, CRect rcItem) = 0;	// Hot
};



////////////////////////////////////////////////////////////////////////////
// CTabSkinTheme

class CTabSkinTheme : public CTabSkin
{
private:
	CTheme&	m_Theme;

public:
	// コンストラクタ
	CTabSkinTheme(CTheme& theme);

private:
	void	_DrawSkinCur(CDCHandle dc, CRect rcItem);
	void	_DrawSkinNone(CDCHandle dc, CRect rcItem);
	void	_DrawSkinSel(CDCHandle dc, CRect rcItem);
};

////////////////////////////////////////////////////////////////////////////
// CTabSkinDefault

class CTabSkinDefault : public CTabSkin
{
private:
	CBitmap		m_TabSkinCur;
	CBitmap		m_TabSkinNone;
	CBitmap		m_TabSkinSel;

public:
	// コンストラクタ
	CTabSkinDefault();

private:
	void	_LoadBitmap();

	void	_DrawSkin(HDC hDC, CRect rcItem, CBitmap& pBmp);

	void	_DrawSkinCur(CDCHandle dc, CRect rcItem);
	void	_DrawSkinNone(CDCHandle dc, CRect rcItem);
	void	_DrawSkinSel(CDCHandle dc, CRect rcItem);
};



////////////////////////////////////////////////////////////////////////////
// CTabItem : タブ１個の情報

class CTabItem
{
public:
	WORD			m_fsState;
	CString			m_strItem;			// タブ名
	CRect			m_rcItem;
	int				m_nImgIndex;		// ImageListのインデックス

	LPITEMIDLIST	m_pidl;				// アイテムＩＤリスト
	CString			m_strFullPath;		// フルパス
	int				m_nSelectedIndex;	// 選択されたアイテムのインデックス
	TRAVELLOG* 		m_pTravelLogBack;
	TRAVELLOG* 		m_pTravelLogFore;

public:
	// コンストラクタ
	CTabItem(LPCTSTR strText = NULL, int nImgIndex = -1, LPITEMIDLIST pidl = NULL, LPCTSTR strFullPath = NULL, BYTE fsState = TCISTATE_ENABLED);

	bool	ModifyState(WORD fsRemove, WORD fsAdd);
#if 0
	void	Update(CDCHandle dc, HIMAGELIST hImgList, bool bAnchorColor, CTabSkin* pTabSkin);

private:
	void	Update_Selected(CDCHandle dc, CTabSkin* pTabSkin, BOOL bHot, BOOL bPressed);
	void	Update_MultiSel(CDCHandle dc, CTabSkin* pTabSkin, BOOL bHot, BOOL bPressed);
	void	Update_NotSel(CDCHandle dc, CTabSkin* pTabSkin, BOOL bHot, BOOL bPressed, CPoint& ptOffset);

	void	_DrawText(CDCHandle dc, CPoint ptOffset, bool bAnchorColor, CTabSkin* pTabSkin);

	void	DrawHot(CDCHandle dc);
	void	DrawPressed(CDCHandle dc);
	void	DrawSelected(CDCHandle dc);
#endif
};


//////////////////////////
// MtlWin.hから
// Locking redraw : 再描写を抑制する
class CLockRedraw {
	CWindow 	m_wnd;

public:
	CLockRedraw(HWND hWnd)
		: m_wnd(hWnd)
	{
		if (m_wnd.m_hWnd) {
			m_wnd.SetRedraw(FALSE);
		}
	}

	~CLockRedraw()
	{
		if (m_wnd.m_hWnd) {
			m_wnd.SetRedraw(TRUE);
			m_wnd.Invalidate();
			m_wnd.UpdateWindow();
		}
	}
};

//////////////////////////


#if 0

/////////////////////////////////////////////////////////////////////////////
// CTabCtrl2 - MTL implementation of Tab Ctrl2

template <class T, class TBase = CWindow, class TWinTraits = CControlWinTraits>
class ATL_NO_VTABLE CTabCtrl2Impl
	: public CWindowImpl< T, TBase, TWinTraits >
	, public CTrackMouseLeave<CTabCtrl2Impl< T, TBase, TWinTraits > >
	, public CThemeImpl<CTabCtrl2Impl< T, TBase, TWinTraits > >
{
public:
	DECLARE_WND_CLASS_EX( NULL, CS_DBLCLKS, -1)

private:
	typedef CTabCtrl2Impl< T, TBase, TWinTraits >	thisClass;

public:
	enum _TabCtrl2DrawConstants {
		s_kcxTextMargin = 7,
		s_kcyTextMargin = 3,
		s_kcxGap		= 2,	// タブ間の隙間
		s_kcyGap		= -2,	// タブの上下間の幅
		s_kSideMargin	= 2,	// 左右の余白
		s_kcxSeparator	= 2,

		s_kcxUpDown 	= 28,
		s_kcyUpDown 	= 14,
	};

private:
	enum {
		_nMaxMenuItemTextLength = 100
	};

	enum _CurrentState {
		none = 0,
		pressed,
		captured,
		hot_by_mouse
	};

public:										// data members
	vector<CTabCtrlItem>	m_items;
	CImageList				m_imgs;

	DWORD					m_dwTabCtrl2ExtendedStyle;

	int						m_nFirstIndexOnSingleLine;	// for single line
	CUpDownCtrl				m_wndUpDown;				//

	CDropDownButtonCtrl		m_wndDropBtn;

protected:
	CSize					m_sizeItem;					// for fixed size

private:
	int						m_nActiveIndex;				// 現在アクティブなインデックス

	_CurrentState			m_nCurrentState;

	int						m_nHotIndex;				// Current hot index
	int						m_nPressedIndex;			// Current capturing index
	CSimpleArray<CPoint>	m_arrSeparators;			// Point of separator

	bool					m_bLockRefreshBandInfo;

	CTabSkin*				m_pTabSkin;
	CToolTipCtrl			m_tip;

	int						m_nDrawStyle;

public:
	// コンストラクタ
	CTabCtrl2Impl()
		: m_dwTabCtrl2ExtendedStyle(TAB2_EX_DEFAULT_BITS)
		, m_nCurrentState(none)
		, m_nActiveIndex(-1)
		, m_nFirstIndexOnSingleLine(0)
		, m_sizeItem(110, 24)
		, m_nHotIndex(-1)
		, m_nPressedIndex(-1)
		, m_bLockRefreshBandInfo(false)
		, m_nDrawStyle(0)
		, m_tabSkin( static_cast<CTheme&>(*this) )
	{
		m_tabSkin.LoadTabSkin();
		SetThemeClassList(L"TAB");
	}

private:
	// Overridables
	void	OnSetCurSel(int nIndex, int nOldIndex) {}

	CString	OnGetToolTipText(int nIndex)
	{
		CString	strItem;

		GetItemText(nIndex, strItem);
		return strItem;
	}

	void	OnDropDownButton() {}

public:


	void	ModifyTabCtrl2ExtendedStyle(DWORD dwRemove, DWORD dwAdd)
	{
		DWORD	dwOldStyle = m_dwTabCtrl2ExtendedStyle;

		m_dwTabCtrl2ExtendedStyle = (m_dwTabCtrl2ExtendedStyle & ~dwRemove) | dwAdd;

		if (dwOldStyle != m_dwTabCtrl2ExtendedStyle) {
			m_nFirstIndexOnSingleLine = 0;
			_UpdateLayout();
		}

		if ( (dwOldStyle & TAB2_EX_ANCHORCOLOR) != (m_dwTabCtrl2ExtendedStyle & TAB2_EX_ANCHORCOLOR) ){
			Invalidate();
		}
	}

	bool	SetItemInactive(int nIndex)
	{
		if ( !IsValidIndex(nIndex) )
			return false;

		if ( m_items[nIndex].ModifyState( 0, TCISTATE_INACTIVE) ) {
			InvalidateRect(m_items[nIndex].m_rcItem, FALSE);
			return true;
		}

		return false;
	}

	// nIndexをアクティブなタブとして表示する
	bool	SetItemActive(int nIndex)
	{
		if ( !IsValidIndex(nIndex) )
			return false;

		if ( m_items[nIndex].ModifyState(TCISTATE_INACTIVE, 0) ) {
			InvalidateRect(m_items[nIndex].m_rcItem, FALSE);
			return true;
		}

		return false;
	}

	bool	SetItemDisabled(int nIndex)
	{
		if ( !IsValidIndex(nIndex) )
			return false;

		if ( m_items[nIndex].ModifyState(TCISTATE_ENABLED, 0) ) {
			InvalidateRect(m_items[nIndex].m_rcItem);
			return true;
		}

		return false;
	}

	bool	SetItemEnabled(int nIndex)
	{
		if ( !IsValidIndex(nIndex) )
			return false;

		if ( m_items[nIndex].ModifyState(0, TCISTATE_ENABLED) ) {
			InvalidateRect(m_items[nIndex].m_rcItem);
			return true;
		}

		return false;
	}


	bool	GetItemText(int nIndex, CString &str)
	{
		if ( !IsValidIndex(nIndex) )
			return false;

		str = m_items[nIndex].m_strItem;

		return true;
	}


	int		GetFirstVisibleIndex()
	{
		if (m_dwTabCtrl2ExtendedStyle & TAB2_EX_MULTILINE) {
			return 0;
		} else {
			CRect	rc;
			GetClientRect(&rc);

			for (int i = 0; i < (int) m_items.size(); ++i) {
				if ( MtlIsCrossRect(m_items[i].m_rcItem, &rc) ) {
					return i;
				}
			}
		}

		return -1;
	}

	int		GetLastVisibleIndex()
	{
		if (m_dwTabCtrl2ExtendedStyle & TAB2_EX_MULTILINE) {
			return (int) m_items.size() - 1;
		} else {
			CRect	rc;
			GetClientRect(&rc);

			for (int i = int(m_items.size()) -1; i >= 0; i--) {
				if ( MtlIsCrossRect(m_items[i].m_rcItem, &rc) ) {
					return i;
				}
			}
		}

		return -1;
	}


	bool	AddItem(const CTabCtrlItem &item)
	{
		_HotItem();

		m_items.push_back(item);
		_UpdateLayout();

		return true;
	}

	/*
	// nIndexのタブをアクティブにする(※この関数ではビューは切り替えない)
	bool	SetCurSel(int nIndex, bool bClicked = false)
	{
		TCTRACE( _T("CTabCtrl2Impl::SetCurSel\n") );
		TCTRACE( _T(" %2d番目のタブが選択されました\n"), nIndex);

		if ( !IsValidIndex(nIndex) ) {
			TCTRACE(_T(" invalid index\n"));
			return false;
		}

		_ResetMultiSelectedItems();					// 複数選択を解除

		int	nOldIndex = GetCurSel();				// 選択される前のタブのインデックスを取得

		if (nOldIndex != -1 && nOldIndex == nIndex) {
			TCTRACE(_T(" no need (同じタブが選択された)\n"));
			return true;
		}

		// clean up
		if ( IsValidIndex(nOldIndex) ) {
			if ( m_items[nOldIndex].ModifyState(TCISTATE_SELECTED | TCISTATE_HOT | TCISTATE_PRESSED | TCISTATE_MSELECTED, 0) ) {
				CRect	rc = m_items[nOldIndex].m_rcItem;
				rc.left	-= s_kSideMargin;
				rc.right+= s_kSideMargin;
				InvalidateRect(rc);
			}
		}

		// new selected item
		if ( m_items[nIndex].ModifyState(TCISTATE_HOT | TCISTATE_PRESSED | TCISTATE_MSELECTED, TCISTATE_SELECTED) ) {
			CRect	rc = m_items[nIndex].m_rcItem;
			rc.left	-= s_kSideMargin;
			rc.right+= s_kSideMargin;
			InvalidateRect(rc);
		}

		m_nActiveIndex = nIndex;		// アクティブにしたタブのインデックスを保存しておく

		_ScrollOpposite(nIndex, bClicked);

		T * pT = static_cast<T *>(this);
		pT->OnSetCurSel(nIndex, nOldIndex);		// 実際にビューを切り替える処理を入れる

		UpdateWindow();
		return true;
	}
	*/


	void	SetDrawStyle(int nStyle)
	{
		m_tabSkin.SetDrawStyle(nStyle);
		m_wndDropBtn.SetDrawStyle(nStyle);
	  #if 0	//+++ uxtheme.dll の関数の呼び出し方を変更.
		if (nStyle == SKN_TAB_STYLE_CLASSIC)
			UxTheme_Wrap::SetWindowTheme(m_wndUpDown.m_hWnd, L" ", L" ");
		else
			UxTheme_Wrap::SetWindowTheme(m_wndUpDown.m_hWnd, NULL, L"SPIN");
	  #else
		CTheme		theme;
		theme.Attach(m_hTheme);
		if (nStyle == SKN_TAB_STYLE_CLASSIC) {
			::SetWindowTheme(m_wndUpDown.m_hWnd, L" ", L" ");
//			theme.SetWindowTheme(m_wndUpDown.m_hWnd, L" ", L" ");
		} else {
			::SetWindowTheme(m_wndUpDown.m_hWnd, NULL, L"SPIN");
//			theme.SetWindowTheme(m_wndUpDown.m_hWnd, NULL, L"SPIN");
		}
	  #endif
	}


	// メッセージマップ
	
	BEGIN_MSG_MAP(CTabCtrl2Impl)
		CHAIN_MSG_MAP(CThemeImpl<CTabCtrl2Impl>)
		COMMAND_ID_HANDLER(COMMAND_ID_DROPBUTTON, OnPushDropButton	)
		MESSAGE_HANDLER( WM_CREATE				, OnCreate				)
		MESSAGE_HANDLER( WM_PAINT				, OnPaint				)
		MESSAGE_HANDLER( WM_WINDOWPOSCHANGING	, OnWindowPosChanging	)
		MESSAGE_HANDLER( WM_LBUTTONDOWN			, OnLButtonDown			)
		MESSAGE_HANDLER( WM_LBUTTONUP			, OnLButtonUp			)
		MESSAGE_HANDLER( WM_ERASEBKGND			, OnEraseBackground		)
		NOTIFY_CODE_HANDLER( UDN_DELTAPOS		, OnDeltaPos	)
		NOTIFY_CODE_HANDLER( TTN_GETDISPINFO	, OnGetDispInfo	)
		CHAIN_MSG_MAP( CTrackMouseLeave<CTabCtrl2Impl> )
		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()
	

private:

	//bool	GetItemRect(int nIndex, CRect &rc)
	//{
	//	if ( !IsValidIndex(nIndex) )
	//		return false;

	//	if (!(m_dwTabCtrl2ExtendedStyle & TAB2_EX_MULTILINE) && nIndex < m_nFirstIndexOnSingleLine)
	//		return false;
	//	rc = m_items[nIndex].m_rcItem;

	//	return true;
	//}




	LRESULT	OnPushDropButton(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL &bHandled)
	{
		T* pT = static_cast<T*>(this);

		pT->OnDropDownButton();

		return 0;
	}


	bool	_MustBeInvalidateOnMultiLine(CSize size)
	{
		CRect	rc;
		GetClientRect(rc);

		if (rc.Width() < size.cx)
			return true;

		int	cxLeft = 0;

		for (int i = 0; i < m_items.size(); ++i) {
			cxLeft = std::max((int)m_items[i].m_rcItem.right, cxLeft);
		}

		if (cxLeft != 0 && cxLeft < size.cx) {
			return false;
		} else {
			return true;
		}
	}


};



//////////////////////////////////////////////////////////////////////////
// CTabCtrl2

class CTabCtrl2 : public CTabCtrl2Impl<CTabCtrl2>
{
public:
	DECLARE_WND_CLASS_EX(_T("MTL_TabCtrl2"), CS_DBLCLKS, COLOR_BTNFACE);
};

#endif

}	// namespace MTL


using namespace MTL;







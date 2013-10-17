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
	TAB2_EX_SUPPORTREBAR	= 0x00000010L,
	TAB2_EX_SCROLLOPPOSITE	= 0x00000020L,
	TAB2_EX_ANCHORCOLOR 	= 0x00000040L,	// タブを表示、未表示で色を変える
	TAB2_EX_MOUSEDOWNSELECT = 0x00000080L,	// タブを左クリックしたときにタブを切り替える
											// ※Drag&Dropでタブを並び変えられない

	TAB2_EX_DEFAULT_BITS	= 
					TAB2_EX_TRANSPARENT 
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
	CString			m_strDrawPath;

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


}	// namespace MTL


using namespace MTL;







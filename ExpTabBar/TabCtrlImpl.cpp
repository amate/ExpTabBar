// TabCtrlImpl.cpp

#include "stdafx.h"
#include "TabCtrlImpl.h"
#include "ExpTabBarOption.h"

////////////////////////////////////////////////////////////////////
// CTabCtrlImpl

CTabCtrlImpl::CTabCtrlImpl()
	: m_nActiveIndex(-1)
	, m_nHotIndex(-1)
	, m_nPressedIndex(-1)
	, m_nFirstIndexOnSingleLine(0)
	, m_bLockRefreshBandInfo(false)
	, m_pTabSkin(NULL)
{
	m_imgs.Create(CXICON, CYICON, ILC_COLOR32 | ILC_MASK, 50, 254);
	WTL::CIcon	hIcon;
	hIcon.LoadIcon(IDI_FOLDER_OPEN);
	m_imgs.AddIcon(hIcon);

	SetThemeClassList(VSCLASS_TAB);
#if 0
	SetTabStyle(SKN_TAB_STYLE_THEME);
	ReloadSkinData();
#endif
}

CTabCtrlImpl::~CTabCtrlImpl()
{
	if (m_imgs.m_hImageList)
		m_imgs.Destroy();		// イメージリストを破棄する

	for (int i = 0; i < GetItemCount(); ++i) {
		// アイテムＩＤリストを解放する
		::CoTaskMemFree(GetItemIDList(i));
		// トラベルログを削除する
		if (GetItem(i).m_pTravelLogBack)
			delete GetItem(i).m_pTravelLogBack;
		if (GetItem(i).m_pTravelLogFore)
			delete GetItem(i).m_pTravelLogFore;
	}
}


void	CTabCtrlImpl::_InitToolTip()
{
	// create a tool tip
	m_tip.Create(m_hWnd);
	ATLASSERT( m_tip.IsWindow() );

	CToolInfo	tinfo(TTF_SUBCLASS, m_hWnd);
	tinfo.hwnd	= m_hWnd;	//\\+ こうしないとこのウィンドウにGetDispInfoが来ない
	m_tip.AddTool(tinfo);
	m_tip.SetMaxTipWidth(SHRT_MAX);
	m_tip.SetDelayTime(TTDT_AUTOPOP, 30 * 1000);
}

// nIndexが有効ならtrue、無効ならfalseを返す
bool	CTabCtrlImpl::_IsValidIndex(int nIndex) const
{
	return ( 0 <= nIndex && nIndex < (int)m_items.size() );
}

// 必要とするタブの高さを返す
int		CTabCtrlImpl::_GetRequiredHeight()
{
	int size = GetItemCount();
	if (size == 0) {
		return GetItemHeight();
	} else {
//		return GetItem(size - 1).m_rcItem.bottom;
		return GetItem(size - 1).m_rcItem.bottom + 1;
	}
}


// タブの大きさを計算する
CRect	CTabCtrlImpl::_MeasureItemRect(const CString &strText)
{
	if (CTabBarConfig::s_bUseFixedSize)
		return CRect(0, 0, CTabBarConfig::s_FixedSize.cx, CTabBarConfig::s_FixedSize.cy);

	CString strTab = MtlCompactString(strText, CTabBarConfig::s_nMaxTextLength);
	// compute size of text - use DrawText with DT_CALCRECT
	int cx = MtlComputeWidthOfText(strTab, m_pTabSkin->GetFontHandle());

	int cxIcon = 0;
	int	cyIcon = 0;
	if (m_imgs.m_hImageList) {
		m_imgs.GetIconSize(cxIcon, cyIcon);
		cyIcon += s_kcyTabIcon * 2;
		cx += cxIcon;
	}

	int cy = m_pTabSkin->GetFontHeight();
	if (cy < 0)
		cy = -cy;
	cy += 2 * s_kcyTextMargin;

	// height of item is the bigger of these two
	cy	= std::max(cy, cyIcon);

	// width is width of text plus a bunch of stuff
	cx += s_kcxTextMargin * 2;	// L/R margin for readability

	return CRect(0, 0, cx, cy);
}

// 現在カーソルが乗っているタブを強調表示する
void	CTabCtrlImpl::_HotItem(int nNewHotIndex)
{
	TCTRACE( _T("_HotItem\n") );

	// clean up
	if ( _IsValidIndex(m_nHotIndex) ) {
		TCTRACE( _T(" clean up - %d\n"), m_nHotIndex);
		CTabItem&	item = GetItem(m_nHotIndex);
		if ( item.ModifyState(TCISTATE_HOT, 0) )
			InvalidateRect(item.m_rcItem);
	}

	m_nHotIndex = nNewHotIndex;

	if ( _IsValidIndex(m_nHotIndex) ) {
		TCTRACE( _T(" hot - %d\n"), m_nHotIndex );
		CTabItem& item = GetItem(m_nHotIndex);

		if ( item.ModifyState(0, TCISTATE_HOT) )
			InvalidateRect(item.m_rcItem);
	}
}


void	CTabCtrlImpl::_PressItem(int nNewPressedIndex)
{
	// clean up prev
	if ( _IsValidIndex(m_nPressedIndex) ) {
		CTabItem&	item = GetItem(m_nPressedIndex);

		if ( item.ModifyState(TCISTATE_PRESSED, 0) )
			InvalidateRect(item.m_rcItem);
	}

	m_nPressedIndex = nNewPressedIndex;

	if ( _IsValidIndex(m_nPressedIndex) ) {
		CTabItem&	item = GetItem(m_nPressedIndex);

		if ( item.ModifyState(0, TCISTATE_PRESSED) )
			InvalidateRect(item.m_rcItem);
	}
}

// 現在複数選択されている中にnIndexがあればtrue
bool	CTabCtrlImpl::_FindIndexFromCurMultiSelected(int nIndex)
{
	CSimpleArray<int>	arrCurMultiSel;
	GetCurMultiSel(arrCurMultiSel, false);

	if (arrCurMultiSel.Find(nIndex) != -1)
		return true;

	return false;
}

// 複数選択されたアイテムを解除する
void	CTabCtrlImpl::_ResetMultiSelectedItems()
{
	for (int i = 0; i < GetItemCount(); ++i) {
		if ( GetItem(i).ModifyState(TCISTATE_MSELECTED, 0) )
			InvalidateRect(GetItem(i).m_rcItem);
	}
}

// 一列表示時に右側にあるボタン類を表示したり隠したりする
void	CTabCtrlImpl::_ShowOrHideUpDownCtrl(const CRect &rcClient)
{
	int	nCount = GetItemCount();

	if (CTabBarConfig::s_bMultiLine || nCount < 1) {
		m_wndDropBtn.ShowWindow(SW_HIDE);
		m_wndUpDown.ShowWindow(SW_HIDE);
		return;
	}

	m_wndUpDown.SetRange( 0, nCount );

	if (m_nFirstIndexOnSingleLine != 0) {
		m_wndDropBtn.ShowWindow(SW_SHOWNORMAL);
		m_wndUpDown.ShowWindow(SW_SHOWNORMAL);
		return;
	} else if (m_items[nCount - 1].m_rcItem.right > rcClient.right) {
		m_wndDropBtn.ShowWindow(SW_SHOWNORMAL);
		m_wndUpDown.ShowWindow(SW_SHOWNORMAL);
		return;
	}

	m_wndDropBtn.ShowWindow(SW_HIDE);
	m_wndUpDown.ShowWindow(SW_HIDE);
}

void	CTabCtrlImpl::_ScrollOpposite(int nNewSel, bool bClicked)
{
	ATLASSERT( _IsValidIndex(nNewSel) );

	if (CTabBarConfig::s_bMultiLine)
		return;

	if (GetItemCount() < 2)
		return;
#if 0
	if ( !(GetTabCtrlExtendedStyle() & TAB2_EX_SCROLLOPPOSITE) )
		return;
#endif
	if (m_wndUpDown.IsWindowVisible() == FALSE)
		return;

	const CRect	&rcItem = GetItem(nNewSel).m_rcItem;

	CRect	rcClient;
	GetClientRect(rcClient);

	if (bClicked) {
		if (m_nFirstIndexOnSingleLine != nNewSel) {
			CRect	rcUpDown;
			m_wndUpDown.GetWindowRect(rcUpDown);
			ScreenToClient(rcUpDown);

			if ( MtlIsCrossRect(rcItem, rcUpDown) )
				ScrollItem();
		}
	} else if (rcItem.IsRectEmpty() || (rcClient & rcItem) != rcItem) {
		m_nFirstIndexOnSingleLine = std::max(0, nNewSel - 1);
		_UpdateSingleLineLayout(m_nFirstIndexOnSingleLine);
		Invalidate();
		_ShowOrHideUpDownCtrl(rcClient);
	}
}


// タブバーが複数行になるときなどに呼ばれる
void	CTabCtrlImpl::_RefreshBandInfo()
{
	if (m_bLockRefreshBandInfo)
		return;

	TCTRACE( _T("CTabCtrlImpl::_RefreshBandInfo\n") );

	HWND		  hWndReBar = GetParent();
	CReBarCtrl	  rebar(hWndReBar);

	static UINT wID = 0;
	if (wID == 0) {
		REBARBANDINFO rb;
		rb.cbSize	= sizeof (REBARBANDINFO);
		rb.fMask	= RBBIM_CHILD | RBBIM_ID;
		int nCount	= rebar.GetBandCount();
		for (int i = 0; i < nCount; ++i) {
			rebar.GetBandInfo(i, &rb);
			if (rb.hwndChild == m_hWnd) {
				wID = rb.wID;
				break;
			}
		}
	}

	int	nIndex = rebar.IdToIndex( wID );
	//		int	nIndex = rebar.IdToIndex(TabBarCtrlID);
	if ( nIndex == -1 ) {
		return;				// 設定できないので帰る
	}

	REBARBANDINFO rbBand;
	rbBand.cbSize = sizeof (REBARBANDINFO);
	rbBand.fMask  = RBBIM_CHILDSIZE;
	rebar.GetBandInfo(nIndex, &rbBand);

	if ( rbBand.cyMinChild != _GetRequiredHeight() ) {
		// Calculate the size of the band
		rbBand.cxMinChild = 0;
		rbBand.cyMinChild = _GetRequiredHeight();

		rebar.SetBandInfo(nIndex, &rbBand);
	}
}

// nFirstIndexがタブバーの一番左にあるタブのインデックスになる
void	CTabCtrlImpl::_UpdateSingleLineLayout(int nFirstIndex, bool bForce)
{
	ATLASSERT( _IsValidIndex(nFirstIndex) );

	m_arrSeparators.RemoveAll();

	// clean invisible items
	int	i;

	for(i = 0; i < nFirstIndex && i < GetItemCount(); ++i) {
		GetItem(i).m_rcItem.SetRectEmpty();
	}

	int	cxOffset = 0;
	for (i = nFirstIndex; i < GetItemCount(); ++i) {
		CTabItem&		item 	= GetItem(i);
		const CRect&	rcSrc	= item.m_rcItem;

		// update the rect
		if (CTabBarConfig::s_bUseFixedSize) {
			item.m_rcItem = CRect(cxOffset, 0, cxOffset + CTabBarConfig::s_FixedSize.cx, CTabBarConfig::s_FixedSize.cy);
		} else {
			item.m_rcItem = _MeasureItemRect(m_items[i].m_strItem) + CPoint(cxOffset, 0);
		}

		cxOffset = cxOffset + item.m_rcItem.Width();

		m_arrSeparators.Add( CPoint(cxOffset, 0) );

		if (bForce || rcSrc != item.m_rcItem) {
			InvalidateRect( _InflateGapWidth(rcSrc) );
			InvalidateRect( _InflateGapWidth(item.m_rcItem) );
		}
	}
}

void	CTabCtrlImpl::_UpdateMultiLineLayout(int nWidth)
{
	TCTRACE(_T("_UpdateMultiLineLayout\n"));

	m_arrSeparators.RemoveAll();

	int	cxOffset = 0;
	int	cyOffset = 0;

	for (int i = 0; i < GetItemCount(); ++i) {	// update the rect
		CTabItem&	item	= GetItem(i);
		const CRect	rcSrc	= item.m_rcItem;

		if (cxOffset == 0) 
			cxOffset = s_kSideMargin;

		if (CTabBarConfig::s_bUseFixedSize) {
			item.m_rcItem = CRect(CPoint(cxOffset, cyOffset), CTabBarConfig::s_FixedSize);
		} else {
			item.m_rcItem = _MeasureItemRect(m_items[i].m_strItem) + CPoint(cxOffset, cyOffset);
		}

		// タブバーの幅以上なら下に移動
		if (i != 0 && item.m_rcItem.right > nWidth - s_kSideMargin) {
			cxOffset = s_kSideMargin;
			cyOffset += GetItemHeight() + s_kcyGap;	// s_kcyGap分下に移動

			if (CTabBarConfig::s_bUseFixedSize) {
				item.m_rcItem = CRect(CPoint(cxOffset, cyOffset), CTabBarConfig::s_FixedSize);
			} else {
				item.m_rcItem = _MeasureItemRect(m_items[i].m_strItem) + CPoint(cxOffset, cyOffset);
			}
		}

		cxOffset = cxOffset + m_items[i].m_rcItem.Width();

		m_arrSeparators.Add( CPoint(cxOffset, cyOffset) );

		if (rcSrc != item.m_rcItem) {
			InvalidateRect( _InflateGapWidth(rcSrc) );
			InvalidateRect( _InflateGapWidth(item.m_rcItem) );
		}
	}
}

void	CTabCtrlImpl::_UpdateLayout()
{
	if (GetItemCount() == 0) {
		m_arrSeparators.RemoveAll();
		Invalidate();

		_RefreshBandInfo();
		return;
	}

	CRect	rc;
	GetClientRect(&rc);

	if (CTabBarConfig::s_bMultiLine) {
		// タブの複数列表示
		_UpdateMultiLineLayout( rc.Width() );
	} else {
		// タブの一列表示
		_UpdateSingleLineLayout(m_nFirstIndexOnSingleLine);
	}

	_ShowOrHideUpDownCtrl(rc);

	_RefreshBandInfo();

	UpdateWindow();
}



const CRect	CTabCtrlImpl::_InflateGapWidth(const CRect &rc) const
{
	int cxSeparatorOffset = (s_kcxGap * 2) + s_kcxSeparator;

	return CRect(rc.left, rc.top, rc.right + cxSeparatorOffset, rc.bottom);
}






void	CTabCtrlImpl::ReloadSkinData()
{
	if (m_pTabSkin) {
		delete m_pTabSkin;
	}

	if (GetTabStyle() & SKN_TAB_STYLE_DEFAULT) {
		m_pTabSkin = new CTabSkinDefault;
	} else {
		m_pTabSkin = new CTabSkinTheme(static_cast<CTheme&>(*this));
	}

	InvalidateRect(NULL, TRUE);
}


void	CTabCtrlImpl::OnTrackMouseMove(UINT nFlags, CPoint pt)
{
	int nIndex = HitTest(pt);
	if ( _IsValidIndex(nIndex) ) {
		// if other button made hot or first hot
		if (m_nHotIndex == -1 || m_nHotIndex != nIndex) {
			m_tip.Activate(FALSE);
			m_tip.Activate(TRUE);

			_HotItem(nIndex);
		}
	} else {
		m_tip.Activate(FALSE);
		_HotItem();	// clean up
	}
}

void	CTabCtrlImpl::OnTrackMouseLeave()
{
	m_tip.SetDelayTime(TTDT_INITIAL, 500);
	_HotItem();		// clean up
}


void	CTabCtrlImpl::DoPaint(CDCHandle dc)
{
	/* タブバーを透明にするため 背景をコピー */
	HWND	hWnd = GetParent();
	CPoint	pt;
	MapWindowPoints(hWnd, &pt, 1);
	pt = ::OffsetWindowOrgEx(dc, pt.x, pt.y, NULL);
	LRESULT lResult = ::SendMessage(hWnd, WM_ERASEBKGND, (WPARAM)dc.m_hDC, 0L);
	::SetWindowOrgEx(dc, 0, 0, NULL);


	CFontHandle	fontOld = dc.SelectFont(m_pTabSkin->GetFontHandle());
	int			modeOld	= dc.SetBkMode(TRANSPARENT);

	int	i = CTabBarConfig::s_bMultiLine ? 0 : m_nFirstIndexOnSingleLine;

	int		nCurIndex = GetCurSel();
	bool	bAnchorColor = false;//(GetTabCtrlExtendedStyle() & TAB2_EX_ANCHORCOLOR) != 0;
	for (; i < GetItemCount(); ++i) {
			// 各タブを描写する
			m_pTabSkin->Update(dc, m_imgs, GetItem(i), bAnchorColor);
	}
	
	if (nCurIndex != -1) {
		// 最後に現在選択されているタブを描写する
		m_pTabSkin->Update(dc, m_imgs, GetItem(nCurIndex), bAnchorColor);
	}
	//		if (m_tabSkin.IsVisible() == FALSE)
	//			_DrawSeparators(dc, lpRect);

	dc.SelectFont(fontOld);
	dc.SetBkMode(modeOld);
}

int		CTabCtrlImpl::GetItemHeight() const
{
	if (CTabBarConfig::s_bUseFixedSize) {
		return CTabBarConfig::s_FixedSize.cy;
	} else {
		int	cxIcon = 0;
		int	cyIcon = 0;

		if (m_imgs.m_hImageList != NULL)
			m_imgs.GetIconSize(cxIcon, cyIcon);

		int	cy = m_pTabSkin->GetFontHeight();

		if (cy < 0)
			cy = -cy;

		cy += s_kcyTextMargin * 2;

		return std::max(cy, cxIcon);
	}
}



// 現在(複数)選択されているタブのインデックスを取得
// bIncludeCurSel == trueで現在アクティブなタブを含める
void	CTabCtrlImpl::GetCurMultiSel(CSimpleArray<int> &arrDest, bool bIncludeCurSel)
{
	arrDest.RemoveAll();

	for (int i = 0; i < GetItemCount(); ++i) {
		if (bIncludeCurSel) {
			if (GetItemState(i) & TCISTATE_SELECTED) {
				arrDest.Add(i);
				continue;
			}
		}

		if (GetItemState(i) & TCISTATE_MSELECTED) {
			arrDest.Add(i);
		}
	}
}


// 複数選択されている部分だけが欲しい場合や
// アクティブなタブも含めて欲しい場合や
// その他の部分が選択されているかをnIndexで判断して返す
void	CTabCtrlImpl::GetCurMultiSelEx(CSimpleArray<int>& arrDest, int nIndex)
{
	GetCurMultiSel(arrDest);

	if (arrDest.Find(nIndex) == -1) {
		// 複数選択されているタブ以外がクリックされた
		arrDest.RemoveAll();
		arrDest.Add(nIndex);
	} else if (nIndex != GetCurSel()) {
		// 複数選択されているタブだけにする
		arrDest.Remove(GetCurSel());
	}
}

bool	CTabCtrlImpl::CanScrollItem(bool bRight) const
{
	if (CTabBarConfig::s_bMultiLine)
		return false;	// can't
	if (bRight) {
		return	m_nFirstIndexOnSingleLine + 1 < GetItemCount();
	} else {
		return	m_nFirstIndexOnSingleLine - 1 >= 0;
	}
}

int		CTabCtrlImpl::ReplaceIcon(int nIndex, HICON hIcon)
{
	int nIconIndex = GetItemImageIndex(nIndex);
	int nImg = m_imgs.ReplaceIcon(nIconIndex, hIcon);
	::DestroyIcon(hIcon);
	m_items[nIndex].m_nImgIndex = nImg;

	InvalidateRect(GetItem(nIndex).m_rcItem, FALSE);
	UpdateWindow();

	return nImg;
}

int		CTabCtrlImpl::AddIcon(HICON hIcon)
{
	int	nIndex = m_imgs.AddIcon(hIcon);
	::DestroyIcon(hIcon);
	return nIndex;
}






int		CTabCtrlImpl::SetCurSel(int nIndex, bool bClicked)
{
	TCTRACE( _T("SetCurSel\n") );

	if ( _IsValidIndex(nIndex) == false ) {
		TCTRACE(_T(" インデックスが無効です : %d\n"), nIndex);
		return -1;
	}

	_ResetMultiSelectedItems();					// 複数選択を解除

	int	nOldIndex = GetCurSel();				// 選択される前のタブのインデックスを取得

	if (nOldIndex != -1 && nOldIndex == nIndex) {
		TCTRACE(_T(" 同じタブが選択された : %d\n"), nIndex);
		return nIndex;
	}

	// clean up
	if ( _IsValidIndex(nOldIndex) ) {
		if ( GetItem(nOldIndex).ModifyState(TCISTATE_SELECTED | TCISTATE_HOT | TCISTATE_PRESSED | TCISTATE_MSELECTED, 0) ) {
			CRect	rc = m_items[nOldIndex].m_rcItem;
			rc.left	-= s_kSideMargin;
			rc.right+= s_kSideMargin;
			InvalidateRect(rc);
		}
	}

	// new selected item
	if ( GetItem(nIndex).ModifyState(TCISTATE_HOT | TCISTATE_PRESSED | TCISTATE_MSELECTED, TCISTATE_SELECTED) ) {
		CRect	rc = m_items[nIndex].m_rcItem;
		rc.left	-= s_kSideMargin;
		rc.right+= s_kSideMargin;
		InvalidateRect(rc);
	}

	m_nActiveIndex = nIndex;		// アクティブにしたタブのインデックスを保存しておく

	_ScrollOpposite(nIndex, bClicked);

	OnSetCurSel(nIndex, nOldIndex);		// 実際にビューを切り替える処理を入れる

	UpdateWindow();

	TCTRACE( _T(" %2d番目のタブが選択されました\n"), nIndex);
	return true;
}


int		CTabCtrlImpl::GetCurSel()
{
	if (m_nActiveIndex == -1) {
		for (int i = 0; i < GetItemCount(); ++i) {
			if (GetItemState(i) & TCISTATE_SELECTED) {
				m_nActiveIndex = i;
				break;
			}
		}
	}

	return m_nActiveIndex;
}


int		CTabCtrlImpl::HitTest(const CPoint& point)
{
	CRect	rc;
	GetClientRect(&rc);
	if (rc.PtInRect(point) == FALSE)
		return -1;

	int	i = CTabBarConfig::s_bMultiLine ? 0 : m_nFirstIndexOnSingleLine;
	int nCount = GetItemCount();
	for (; i < nCount; ++i) {
		if (m_items[i].m_rcItem.PtInRect(point)){
			return i;
		}
	}

	return -1;
}



int		CTabCtrlImpl::InsertItem(int nIndex, const CTabItem &item)
{
	TCTRACE( _T("InsertItem : %2d\n"), nIndex );

	_HotItem();			// clean up

	if ( _IsValidIndex(nIndex) == false ) {
		nIndex = GetItemCount();	// nIndexが無効なら最後に追加
		m_items.push_back(item);
		TCTRACE( _T(" 最後に追加されました\n") );
	} else {
		m_items.insert(m_items.begin() + nIndex, item);
		m_nActiveIndex = -1;
	}

	_UpdateLayout();

	return nIndex;
}

int		CTabCtrlImpl::AddItem(const CTabItem &item)
{
	return InsertItem(-1, item);
}


void	CTabCtrlImpl::DeleteItem(int nIndex, bool bMoveNow)
{
	if ( _IsValidIndex(nIndex) == false )
		return;

	_HotItem();

//	InvalidateRect( _InflateGapWidth(m_items[nIndex].m_rcItem) );
	InvalidateRect(GetItem(nIndex).m_rcItem);

	if ( nIndex < m_nFirstIndexOnSingleLine || nIndex + 1 == GetItemCount() ) {
		--m_nFirstIndexOnSingleLine;

		if (m_nFirstIndexOnSingleLine < 0)
			m_nFirstIndexOnSingleLine = 0;
	}

	// タブが移動しただけならアイテムＩＤリストを解放しない
	if (bMoveNow == false) {
		// アイテムＩＤリストを解放する
		::CoTaskMemFree(GetItemIDList(nIndex));
		
		int nRemovedIndex = GetItemImageIndex(nIndex);
		if (nRemovedIndex != -1) {
			// アイコンを解放する
			m_imgs.Remove(nRemovedIndex);
			for (int i = 0; i < GetItemCount(); ++i) {
				int nImgIndex = GetItemImageIndex(i);
				if (nImgIndex > nRemovedIndex) {
					// 削除されたイメージのインデックスより上だったので１個ずらす
					--m_items[i].m_nImgIndex;
				}
			}
		}
		
		// トラベルログを削除する
		if (GetItem(nIndex).m_pTravelLogBack)
			delete GetItem(nIndex).m_pTravelLogBack;
		if (GetItem(nIndex).m_pTravelLogFore)
			delete GetItem(nIndex).m_pTravelLogFore;
	}

	m_items.erase(m_items.begin() + nIndex);

	m_nActiveIndex = -1;

	_UpdateLayout();
}

void	CTabCtrlImpl::DeleteItems(const CSimpleArray<int> &arrSrcs)
{
	// 後ろから削除する
	for (int i = arrSrcs.GetSize() - 1; i >= 0; --i) {
		int	nIndex = arrSrcs[i];

		if ( _IsValidIndex(nIndex) == false )
			continue;

		DeleteItem(nIndex);
	}
}


// ドラッグした(複数選択された)タブをnDestIndexに移動する
// ＿＿＿＿＿＿＿＿
// |.0.|..1..|..2.|	←タブのインデックス
// ↑￣↑￣￣↑￣↑
// 0   1     2   3  ←nDestIndex
// 上のように挿入される
bool	CTabCtrlImpl::MoveItems(int nDestIndex, const CSimpleArray<int> &arrSrcs)
{
	if (arrSrcs.GetSize() <= 0)
		return false;

	int	i = 0;

	if (arrSrcs.GetSize() == 1) {
		if (nDestIndex == arrSrcs[0] || nDestIndex == arrSrcs[0] + 1)
			return true;	// 自分のタブの左右にDropされたので帰る

		if (_IsValidIndex(nDestIndex) == false && arrSrcs[0] == GetItemCount() -1)
			return true;
	}

	CLockRedraw	lock(m_hWnd);
	m_bLockRefreshBandInfo = true;

	std::vector<CTabItem>	temp;

	for (i = 0; i < arrSrcs.GetSize(); ++i) {
		int	nIndex = arrSrcs[i];

		if ( !_IsValidIndex(nIndex) )
			continue;

		temp.push_back(m_items[nIndex]);
	}

	/* アイテムを削除する処理 */
	// eraseするとインデックスが詰められるので後ろから消していく
	for (i = arrSrcs.GetSize() - 1; i >= 0; --i) {
		int	nIndex = arrSrcs[i];

		if ( !_IsValidIndex(nIndex) )
			continue;

		if (nDestIndex > nIndex) {
			// タブが右に移動するとき
			// (itemを削除するので、その分挿入位置をマイナスする)
			--nDestIndex;
		}
		DeleteItem(nIndex, true);
	}

	// add all
	if ( !_IsValidIndex(nDestIndex) ) {
		nDestIndex = GetItemCount();	// おかしな値だったので最後に挿入するようにする
	}

	/* アイテムを挿入する処理 */
	for (i = 0; i < (int) temp.size(); ++i) {
		InsertItem(nDestIndex, temp[i]);
		++nDestIndex;
	}

	// アクティブなタブのインデックスが変わったので初期化しておく
	m_nActiveIndex = -1;

	m_bLockRefreshBandInfo = false;
	_RefreshBandInfo();

	return true;
}

// bRight == trueで右にスクロールする
bool	CTabCtrlImpl::ScrollItem(bool bRight)
{
	if (CTabBarConfig::s_bMultiLine)
		return false;		// do nothing

	if (bRight) {
		int 	nCount = GetItemCount();
		if ( m_nFirstIndexOnSingleLine + 1 < nCount ) {
			++m_nFirstIndexOnSingleLine;
			_UpdateSingleLineLayout(m_nFirstIndexOnSingleLine);
			Invalidate();
		} else {
			return false;
		}
	} else {
		if (m_nFirstIndexOnSingleLine - 1 >= 0) {
			--m_nFirstIndexOnSingleLine;
			_UpdateSingleLineLayout(m_nFirstIndexOnSingleLine);
			Invalidate();
		} else {
			return false;
		}
	}

	CRect	rcClient;
	GetClientRect(&rcClient);

	_ShowOrHideUpDownCtrl(rcClient);
	return true;
}



//////////////////////
// CTabItem

void	CTabCtrlImpl::SetItemText(int nIndex, LPCTSTR strText)
{
	ATLASSERT(_IsValidIndex(nIndex));

	CString strTab = strText;
	//strTab = MtlCompactString(strTab, CTabBarConfig::s_nMaxTextLength);
	if (m_items[nIndex].m_strItem == strTab)
		return ;

	m_items[nIndex].m_strItem = strTab;

	if (CTabBarConfig::s_bUseFixedSize) {
		InvalidateRect(m_items[nIndex].m_rcItem);
		UpdateWindow();
	} else {
		InvalidateRect(m_items[nIndex].m_rcItem);	// even if layout will not be changed
		_UpdateLayout();							//_UpdateItems(nIndex);
	}
}

void	CTabCtrlImpl::SetItemImageIndex(int nIndex, int nImgIndex)
{
	CTabItem& item = GetItem(nIndex);

	item.m_nImgIndex = nImgIndex;

	InvalidateRect(item.m_rcItem);
	UpdateWindow();
}

void	CTabCtrlImpl::SetItemIDList(int nIndex, LPITEMIDLIST pidl)
{
	if(GetItemIDList(nIndex)) {
		::CoTaskMemFree(GetItemIDList(nIndex));
	}
	m_items.at(nIndex).m_pidl = pidl;
}


int 	CTabCtrlImpl::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	LRESULT	lRet = DefWindowProc();

	m_wndUpDown.Create(m_hWnd, CRect(0, 0, s_kcxUpDown, s_kcyUpDown), NULL, WS_CHILD | WS_OVERLAPPED | UDS_HORZ);

	m_wndDropBtn.Create(m_hWnd, CRect(0, 0, s_kcxUpDown, s_kcxUpDown), COMMAND_ID_DROPBUTTON);
	m_wndDropBtn.SetWindowPos(m_wndUpDown.m_hWnd, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);

	_InitToolTip();

	// default size
	MoveWindow(0, 0, 200, _GetRequiredHeight(), FALSE);

	return static_cast<int>(lRet);
}


void	CTabCtrlImpl::OnDestroy()
{
	if (m_wndUpDown.IsWindow())
		m_wndUpDown.DestroyWindow();
	if (m_wndDropBtn.IsWindow())
		m_wndDropBtn.DestroyWindow();
}


void	CTabCtrlImpl::OnWindowPosChanging(LPWINDOWPOS lpWndPos)
{
	SetMsgHandled(FALSE);

	CSize	size(lpWndPos->cx, lpWndPos->cy);
	m_wndUpDown.MoveWindow(size.cx - s_kcyUpDown * 3, size.cy - s_kcyUpDown, s_kcxUpDown, s_kcxUpDown, s_kcyUpDown);
	m_wndDropBtn.MoveWindow(size.cx - s_kcyUpDown, size.cy - s_kcyUpDown, s_kcyUpDown, s_kcyUpDown);

	CRect	rc(0, 0, s_kcyUpDown, s_kcyUpDown);
	m_wndDropBtn.InvalidateRect(&rc, TRUE);
	_ShowOrHideUpDownCtrl( CRect(0, 0, size.cx, size.cy) );

	if (CTabBarConfig::s_bMultiLine) {
		_UpdateMultiLineLayout(size.cx);

		_RefreshBandInfo();
	}
}

// 左クリックDown
void	CTabCtrlImpl::OnLButtonDown(UINT nFlags, CPoint point)
{
	TCTRACE(_T("CTabCtrlImpl::OnLButtonDown\n"));

	int	nIndex = HitTest(point);

	if (nIndex != -1) {
		ATLASSERT( _IsValidIndex(nIndex) );

		// Ctrl+クリックされていれば
		if ( (nFlags & MK_CONTROL) && GetCurSel() != nIndex ) {
			if ( _FindIndexFromCurMultiSelected(nIndex) == false ) {
				// 複数選択中の中になかったので複数選択状態にする
				if ( GetItem(nIndex).ModifyState(TCISTATE_SELECTED, TCISTATE_MSELECTED) ) {
					InvalidateRect(GetItem(nIndex).m_rcItem);
				}
			} else {
				// 複数選択中の中にあったので複数選択状態を解除する
				if ( m_items[nIndex].ModifyState(TCISTATE_MSELECTED, 0) ) {
					InvalidateRect(GetItem(nIndex).m_rcItem);
				}
			}
		} else {
			_PressItem(nIndex);
			SetCapture();
		}
	}
}

// 左クリックUp
void	CTabCtrlImpl::OnLButtonUp(UINT nFlags, CPoint point)
{
	TCTRACE(_T("CTabCtrlImpl::OnLButtonUp\n"));

	if (GetCapture() == m_hWnd) {
		ReleaseCapture();

		int	nIndex = HitTest(point);

		if (nIndex != -1 && nIndex == m_nPressedIndex) {
			ATLASSERT( _IsValidIndex(nIndex) );
			_PressItem();					// always clean up pressed flag
			SetCurSel(nIndex, true);
			NMHDR	nmhdr = { m_hWnd, GetDlgCtrlID(), TCN_SELCHANGE };
			::SendMessage(GetParent(), WM_NOTIFY, (WPARAM) GetDlgCtrlID(), (LPARAM) &nmhdr);
			TCTRACE(_T(" タブが切り替わった\n"));
		} else {
			_PressItem();					// alwaisy clean up pressed flag
		}
	}
}


LRESULT CTabCtrlImpl::OnDeltaPos(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
	LPNMUPDOWN lpnmud = (LPNMUPDOWN) pnmh;

	ScrollItem(lpnmud->iDelta > 0);

	return 0;
}

LRESULT CTabCtrlImpl::OnGetDispInfo(int idCtrl, LPNMHDR pnmh, BOOL& bHandled)
{
	LPNMTTDISPINFO pDispInfo = (LPNMTTDISPINFO) pnmh;
	CPoint	pt;
	::GetCursorPos(&pt);
	ScreenToClient(&pt);
	int		nIndex = HitTest(pt);

	if ( _IsValidIndex(nIndex) ) {
		m_strTooltip = m_items[nIndex].m_strFullPath;
		if (m_strTooltip.Left(2) == _T("::"))
			m_strTooltip = m_items[nIndex].m_strItem;
		pDispInfo->lpszText = m_strTooltip.GetBuffer();
	} else {
		pDispInfo->szText[0] = _T('\0');
	}

	m_tip.SetDelayTime( TTDT_INITIAL, m_tip.GetDelayTime(TTDT_RESHOW) );

	return 0;
}















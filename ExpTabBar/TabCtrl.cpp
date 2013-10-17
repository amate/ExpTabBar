// TabCtrl.cpp

#include "stdafx.h"
#include "TabCtrl.h"
//#include "TabBarWindow.h"
#include "Misc.h"
#include "ExpTabBarOption.h"


namespace MTL {


///////////////////////////////////////////////////////////////////////////////////
// CTabSkin

void	CTabSkin::Update(CDCHandle dc, HIMAGELIST hImgList, const CTabItem& item, bool bAnchorColor)
{
	if (item.m_fsState & TCISTATE_HIDDEN)
		return;

	int	cxIcon			= 0;
	int cyIcon			= 0;
	int	cxIconOffset	= 0;

	if (item.m_nImgIndex != -1) {
		::ImageList_GetIconSize(hImgList, &cxIcon, &cyIcon);
		cxIconOffset += cxIcon + s_kcxIconGap + 3;
	}

	bool	bHot		= (item.m_fsState & TCISTATE_HOT) != 0;
	bool	bPressed	= (item.m_fsState & TCISTATE_PRESSED) != 0;

	CPoint	ptOffset(0, 0);

	if (item.m_fsState & TCISTATE_SELECTED) {				// �I�����ꂽ
		_DrawSkinCur(dc, item.m_rcItem);
//		Update_Selected(dc, pTabSkin, bHot, bPressed);
//		ptOffset += CPoint(2, 2);	// �^�u�ɏ�����镶�����኱������
	} else if (item.m_fsState & TCISTATE_MSELECTED) {		// �����I�����ꂽ
		_DrawSkinSel(dc, item.m_rcItem);
//		Update_MultiSel(dc, pTabSkin, bHot, bPressed);
		ptOffset += CPoint(2, 4);	// �^�u�ɏ�����镶�����኱������
	} else {											// �I������Ă��Ȃ�
		if (bHot == true && bPressed == false) {
			_DrawSkinSel(dc, item.m_rcItem);
		} else {
			_DrawSkinNone(dc, item.m_rcItem);
		}
//		Update_NotSel(dc, pTabSkin, bHot, bPressed, ptOffset);
		ptOffset += CPoint(0, 2);
	}

	_DrawText(dc, ptOffset + CPoint(cxIconOffset, 0), item, bAnchorColor);

	if (item.m_nImgIndex != -1) {
		::ImageList_Draw(hImgList, item.m_nImgIndex, dc, 
			item.m_rcItem.left + ptOffset.x + s_kcxIconGap,
			item.m_rcItem.top + ((item.m_rcItem.Height() - cyIcon) / 2) + (ptOffset.y / 2), ILD_TRANSPARENT);
	}

	if (item.m_fsState & TCISTATE_NAVIGATELOCK) {
		CPoint	pt(item.m_rcItem.left + ptOffset.x + s_kcxIconGap + 6
				 , item.m_rcItem.top + ((item.m_rcItem.Height() - cyIcon) / 2) + (ptOffset.y / 2) + 4);
		m_imgLock.Draw(dc, 0, pt, ILD_TRANSPARENT);

	}
}

int		CTabSkin::GetFontHeight() const
{
	LOGFONT	lf = { 0 };
	m_font.GetLogFont(lf);

	return lf.lfHeight;
}



void	CTabSkin::_GetSystemFontSetting(LOGFONT* plf)
{
	// refresh our font
	NONCLIENTMETRICS info = {0};

	info.cbSize = sizeof (info);
	::SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof (info), &info, 0);

	*plf = info.lfMenuFont;
}

void	CTabSkin::_LoadTabTextSetting()
{
	CString	strTabSkinIni = Misc::GetExeDirectory() + _T("TabSkin.ini");

	// �V�X�e������̒l��ݒ�
	LOGFONT	lf = { 0 };
	_GetSystemFontSetting(&lf);

	m_colText		= ::GetSysColor(COLOR_BTNTEXT);
	m_colActive		= RGB( 128, 0, 128 );
	m_colInActive	= RGB(   0, 0, 255 );
	m_colDisable	= ::GetSysColor(COLOR_3DSHADOW);
	m_colDisableHi	= ::GetSysColor(COLOR_3DHILIGHT);
	m_colBtnFace	= ::GetSysColor(COLOR_BTNFACE);
	m_colBtnHi		= ::GetSysColor(COLOR_BTNHILIGHT);

	// �t�@�C���̑��݃`�F�b�N
	if (::GetFileAttributes(strTabSkinIni) == 0xFFFFFFFF) {
			strTabSkinIni = _GetSkinDir() + _T("TabSkin.ini");
	}

	// .ini����ǂݍ���
	CIniFileRead pr( strTabSkinIni, _T("Color") );
	_QueryColorString( pr, m_colText		, _T("Text") );
	_QueryColorString( pr, m_colActive		, _T("Active") );
	_QueryColorString( pr, m_colInActive	, _T("InActive") );
	_QueryColorString( pr, m_colDisable		, _T("Disable") );
	_QueryColorString( pr, m_colDisableHi	, _T("DisableHilight") );
	_QueryColorString( pr, m_colBtnFace 	, _T("ButtonFace") );
	_QueryColorString( pr, m_colBtnHi		, _T("ButtonHilight") );

	pr.ChangeSectionName(_T("Font"));
	lstrcpy(lf.lfFaceName, 	pr.GetString(_T("lfFaceName") 	, lf.lfFaceName));
	HDC hDC = ::GetDC(NULL);
	lf.lfHeight		= -MulDiv(pr.GetValue (_T("lfHeight"), lf.lfHeight), GetDeviceCaps(hDC, LOGPIXELSY), 72);
	lf.lfItalic		= (BYTE)pr.GetValue (_T("lfItalic")		, lf.lfItalic);
	lf.lfUnderline	= (BYTE)pr.GetValue (_T("lfUnderline")	, lf.lfUnderline);
	lf.lfStrikeOut	= (BYTE)pr.GetValue (_T("lfStrikeOut")	, lf.lfStrikeOut);
	lf.lfWeight 	= FW_NORMAL;
	lf.lfCharSet 	= DEFAULT_CHARSET;
	if (m_font.IsNull() == false) {
		m_font.DeleteObject();
	}
	m_font.CreateFontIndirect(&lf);
}

void	CTabSkin::_LoadImage()
{
	m_imgLock.Create(CXICON, CYICON, ILC_COLOR32 | ILC_MASK, 1, 1);

	CString	strTabSkinDir = _GetSkinDir();
	CString strNavigateLock = strTabSkinDir;
	strNavigateLock.Append(_T("NavigateLock.ico"));
//	ATLASSERT(FALSE);
//	HICON hIcon = AtlLoadIconImage(strNavigateLock.GetBuffer(0), LR_LOADFROMFILE/* | LR_LOADTRANSPARENT*/, CXICON, CYICON);
	HICON hIcon = AtlLoadIcon(IDI_NAVIGATELOCK);
	m_imgLock.AddIcon(hIcon);
}


void	CTabSkin::_DrawText(CDCHandle dc, CPoint ptOffset, const CTabItem& item, bool bAnchorColor)
{
	COLORREF	clr;

	if ( !(item.m_fsState & TCISTATE_ENABLED) )
		clr = m_colDisable;
	else if (item.m_fsState & TCISTATE_INACTIVE)
		clr = bAnchorColor ? m_colInActive	: m_colText;
	else
		clr = bAnchorColor ? m_colActive	: m_colText;

	COLORREF	clrOld = dc.SetTextColor(clr);

	CRect		rcBtn(item.m_rcItem.left + ptOffset.x, item.m_rcItem.top + ptOffset.y, 
					  item.m_rcItem.right, item.m_rcItem.bottom);
	rcBtn.DeflateRect(2, 0);

	UINT	uFormat;
	//int		nWidth	= MtlComputeWidthOfText(item.m_strItem, dc.GetCurrentFont());
	CString strTab;
	if (item.m_strDrawPath.GetLength() > 0) {
		strTab = item.m_strDrawPath;
	} else {
		strTab = MtlCompactString(item.m_strItem, CTabBarConfig::s_nMaxTextLength);
	}
	if ( true /*nWidth > rcBtn.Width()*/ ) {
		uFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_LEFT | DT_END_ELLIPSIS;
	} else {
		uFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_CENTER | DT_NOCLIP;
	}

	if ( !(item.m_fsState & TCISTATE_ENABLED) ) {
		CRect		rcDisabled	= rcBtn + CPoint(1, 1);
		COLORREF	clrOld2		= dc.SetTextColor(m_colDisableHi);
		dc.DrawText(strTab/*item.m_strItem*/, -1, rcDisabled, uFormat);
		dc.SetTextColor(clrOld2);
	}

	// �^�u��ɕ����������
	dc.DrawText(strTab/*item.m_strItem*/, -1, rcBtn, uFormat);
	dc.SetTextColor(clrOld);
}




////////////////////////////////////////////////////////////////////////////
// CTabSkinTheme

CTabSkinTheme::CTabSkinTheme(CTheme& theme)
	: m_Theme(theme)
{
}


void	CTabSkinTheme::_DrawSkinCur(CDCHandle dc, CRect rcItem)
{
	rcItem.left -= 2;
	rcItem.right+= 2;
	m_Theme.DrawThemeBackground(dc, TABP_TABITEM, TIS_SELECTED, rcItem);
}

void	CTabSkinTheme::_DrawSkinNone(CDCHandle dc, CRect rcItem)
{
	rcItem.top += 2;
	m_Theme.DrawThemeBackground(dc, TABP_TABITEM, TIS_NORMAL, rcItem);
}

void	CTabSkinTheme::_DrawSkinSel(CDCHandle dc, CRect rcItem)
{
	rcItem.top += 2;
	m_Theme.DrawThemeBackground(dc, TABP_TABITEM, TIS_HOT, rcItem);
}



////////////////////////////////////////////////////////////////////////
// CTabSkinDefault

CTabSkinDefault::CTabSkinDefault()
{
	_LoadBitmap();
}


void	CTabSkinDefault::_LoadBitmap()
{
	// �������[�h����Ă����Ȃ�����Ă����i�O�̂�������j
	if (m_TabSkinCur.m_hBitmap)		m_TabSkinCur.DeleteObject();
	if (m_TabSkinNone.m_hBitmap)	m_TabSkinNone.DeleteObject();
	if (m_TabSkinSel.m_hBitmap)		m_TabSkinSel.DeleteObject();

	CString	strTabSkinDir = _GetSkinDir();	// �X�L����u���Ă���t�H���_���擾
	CString strTabSkinCur;
	CString strTabSkinNone;
	CString strTabSkinSel;

	strTabSkinCur.Format (_T("%sTabSkinCur.bmp")	, strTabSkinDir);
	strTabSkinNone.Format(_T("%sTabSkinNone.bmp")	, strTabSkinDir);
	strTabSkinSel.Format (_T("%sTabSkinSel.bmp")	, strTabSkinDir);

	// �^�u�E�X�L��
	m_TabSkinCur.Attach	( AtlLoadBitmapImage(strTabSkinCur.GetBuffer(0)	, LR_LOADFROMFILE) );
	m_TabSkinNone.Attach( AtlLoadBitmapImage(strTabSkinNone.GetBuffer(0), LR_LOADFROMFILE) );
	m_TabSkinSel.Attach	( AtlLoadBitmapImage(strTabSkinSel.GetBuffer(0)	, LR_LOADFROMFILE) );

	if (m_TabSkinCur.m_hBitmap	== NULL) throw strTabSkinCur;
	if (m_TabSkinNone.m_hBitmap == NULL) throw strTabSkinNone;
	if (m_TabSkinSel.m_hBitmap	== NULL) throw strTabSkinSel;
}


void	CTabSkinDefault::_DrawSkin(HDC hDC, CRect rcItem, CBitmap& pBmp)
{
	CSize	szBmp;
	pBmp.GetSize(szBmp);
	HDC		hdcCompatible	= ::CreateCompatibleDC(hDC);
	HBITMAP	hBitmapOld		= (HBITMAP)SelectObject(hdcCompatible, pBmp.m_hBitmap);

	if ( rcItem.Width() <= szBmp.cx ) {
		// All
		::StretchBlt(hDC, rcItem.left, rcItem.top, rcItem.Width(), rcItem.Height(), hdcCompatible, 0, 0, szBmp.cx, szBmp.cy, SRCCOPY);

		int	nWidthLR = rcItem.Width() / 3;
		if ( nWidthLR < szBmp.cx ) {
			// Left
			::StretchBlt(hDC, rcItem.left, rcItem.top, nWidthLR, rcItem.Height(), hdcCompatible, 0, 0, nWidthLR, szBmp.cy, SRCCOPY);

			// Right
			::StretchBlt(hDC, rcItem.right - nWidthLR, rcItem.top, nWidthLR, rcItem.Height(), hdcCompatible, szBmp.cx - nWidthLR, 0, nWidthLR, szBmp.cy, SRCCOPY);
		}
	} else {
		int	nWidthLR = szBmp.cy / 3;
		// Left
		::StretchBlt(hDC, rcItem.left, rcItem.top, nWidthLR, rcItem.Height(), hdcCompatible, 0, 0, nWidthLR, szBmp.cy, SRCCOPY);

		// Right
		::StretchBlt(hDC, rcItem.right - nWidthLR, rcItem.top, nWidthLR, rcItem.Height(), hdcCompatible, szBmp.cx - nWidthLR, 0, nWidthLR, szBmp.cy, SRCCOPY);

		// Mid
		::StretchBlt(hDC, rcItem.left + nWidthLR, rcItem.top, rcItem.Width() - (nWidthLR * 2), rcItem.Height(), hdcCompatible, nWidthLR, 0, nWidthLR, szBmp.cy, SRCCOPY);
	}

	::SelectObject(hdcCompatible, hBitmapOld);
	::DeleteObject(hdcCompatible);
}


void	CTabSkinDefault::_DrawSkinCur(CDCHandle dc, CRect rcItem)
{
	_DrawSkin(dc, rcItem, m_TabSkinCur);
}

void	CTabSkinDefault::_DrawSkinNone(CDCHandle dc, CRect rcItem)
{
	_DrawSkin(dc, rcItem, m_TabSkinNone);
}

void	CTabSkinDefault::_DrawSkinSel(CDCHandle dc, CRect rcItem)
{
	_DrawSkin(dc, rcItem, m_TabSkinSel);
}





/////////////////////////////////////////////////////////////////////////////
// CTabItem

// �R���X�g���N�^
CTabItem::CTabItem(LPCTSTR strText, int nImgIndex, LPITEMIDLIST pidl, LPCTSTR strFullPath, BYTE fsState)
	: m_strItem(strText)
	, m_pidl(pidl)
	, m_nImgIndex(nImgIndex)
	, m_nSelectedIndex(-1)
	, m_fsState(fsState)
	, m_pTravelLogBack(NULL)
	, m_pTravelLogFore(NULL)
{
}


bool	CTabItem::ModifyState(WORD fsRemove, WORD fsAdd)
{
	WORD	fsStateOld = m_fsState;

	m_fsState = (m_fsState & ~fsRemove) | fsAdd;

	if ( (fsStateOld & TCISTATE_SELECTED || fsStateOld & TCISTATE_MSELECTED)
		&& (fsRemove == TCISTATE_HOT || fsAdd == TCISTATE_HOT) )
		return false;
	else
		return m_fsState != fsStateOld;
}
#if 0

void	CTabCtrlItem::Update(CDCHandle dc, HIMAGELIST hImgList, bool bAnchorColor, CTabSkin *pTabSkin)
{
	if (m_fsState & TCISTATE_HIDDEN)
		return;

	int	cxIcon			= 0;
	int cyIcon			= 0;
	int	cxIconOffset	= 0;

	if (m_nImgIndex != -1) {
		::ImageList_GetIconSize(hImgList, &cxIcon, &cyIcon);
		cxIconOffset += cxIcon + s_kcxIconGap + 1;
	}

	bool	bHot		= (m_fsState & TCISTATE_HOT) != 0;
	bool	bPressed	= (m_fsState & TCISTATE_PRESSED) != 0;

	CPoint	ptOffset(0, 0);

	if (m_fsState & TCISTATE_SELECTED) {				// �I�����ꂽ
		pTabSkin->DrawSkinCur();
//		Update_Selected(dc, pTabSkin, bHot, bPressed);
//		ptOffset += CPoint(2, 2);	// �^�u�ɏ�����镶�����኱������
	} else if (m_fsState & TCISTATE_MSELECTED) {		// �����I�����ꂽ
		pTabSkin->DrawSkinCur();
//		Update_MultiSel(dc, pTabSkin, bHot, bPressed);
		ptOffset += CPoint(2, 2);	// �^�u�ɏ�����镶�����኱������
	} else {											// �I������Ă��Ȃ�
		if (bHot == true && bPressed == false) {
			pTabSkin->DrawSkinSel();
		} else {
			pTabSkin->DrawSkinNone();
		}
//		Update_NotSel(dc, pTabSkin, bHot, bPressed, ptOffset);
		ptOffset += CPoint(0, 2);
	}

	_DrawText(dc, ptOffset + CPoint(cxIconOffset, 0), bAnchorColor, pTabSkin);

	if (m_nImgIndex != -1) {
		::ImageList_Draw(hImgList, m_nImgIndex, dc, 
			m_rcItem.left + ptOffset.x + s_kcxIconGap,
			m_rcItem.top + ((m_rcItem.Height() - cyIcon) / 2) + (ptOffset.y / 2), ILD_TRANSPARENT);
	}
}



void	CTabCtrlItem::Update_Selected(CDCHandle dc, CTabSkin *pTabSkin, BOOL bHot, BOOL bPressed)
{
	if (bHot && bPressed) {
		// �I�����
		if (pTabSkin->IsVisible() == FALSE) {
			// �X�L���Ȃ�
			COLORREF	crTxt	= dc.SetTextColor(pTabSkin->m_colBtnFace);
			COLORREF	crBk	= dc.SetBkColor(pTabSkin->m_colBtnHi);
			CBrush		hbr( CDCHandle::GetHalftoneBrush() );
			dc.SetBrushOrg(m_rcItem.left, m_rcItem.top);
			dc.FillRect(m_rcItem, hbr);
			dc.SetTextColor(crTxt);
			dc.SetBkColor(crBk);
			dc.DrawEdge(m_rcItem, BDR_RAISEDINNER, BF_RECT);
		} else {
			// �X�L���g�p
			pTabSkin->DrawSkinCur(dc.m_hDC, m_rcItem);
		}
	} else {
		if (pTabSkin->IsVisible() == FALSE) {
			// �X�L���Ȃ�
			COLORREF	crTxt	= dc.SetTextColor(pTabSkin->m_colBtnFace);
			COLORREF	crBk	= dc.SetBkColor(pTabSkin->m_colBtnHi);
			CBrush		hbr( CDCHandle::GetHalftoneBrush() );
			dc.SetBrushOrg(m_rcItem.left, m_rcItem.top);
			dc.FillRect(m_rcItem, hbr);
			dc.SetTextColor(crTxt);
			dc.SetBkColor(crBk);
			dc.DrawEdge(m_rcItem, EDGE_SUNKEN, BF_RECT);
		} else {
			pTabSkin->DrawSkinCur(dc.m_hDC, m_rcItem);
		}
	}
}

void	CTabCtrlItem::Update_MultiSel(CDCHandle dc, CTabSkin *pTabSkin, BOOL bHot, BOOL bPressed)
{
	if (pTabSkin->IsVisible() == FALSE) {
		// �X�L���Ȃ�
		if (bHot && bPressed) {
			dc.DrawEdge(m_rcItem, BDR_RAISEDINNER, BF_RECT);
		} else {
			dc.DrawEdge(m_rcItem, EDGE_SUNKEN, BF_RECT);
		}
	} else {
		// �X�L���g�p
		pTabSkin->DrawSkinSel(dc.m_hDC, m_rcItem);
	}
}

void	CTabCtrlItem::Update_NotSel(CDCHandle dc, CTabSkin* pTabSkin, BOOL bHot, BOOL bPressed, CPoint& ptOffset)
{
	if (pTabSkin->IsVisible() == FALSE) {
		// �X�L���Ȃ�
		if (bHot && bPressed) {
			dc.DrawEdge(m_rcItem, BDR_RAISEDINNER, BF_RECT);
			ptOffset += CPoint(2, 2);
		} else if (bHot && !bPressed) {
			dc.DrawEdge(m_rcItem, BDR_RAISEDINNER, BF_RECT);
		}
	} else {
		// �X�L���g�p
		if (bHot && !bPressed) {
			pTabSkin->DrawSkinSel(dc.m_hDC, m_rcItem);
		} else {
			pTabSkin->DrawSkinNone(dc.m_hDC, m_rcItem);
		}
	}
}

void	CTabCtrlItem::_DrawText(CDCHandle dc, CPoint ptOffset, bool bAnchorColor, CTabSkin *pTabSkin)
{
	COLORREF	clr;

	if ( !(m_fsState & TCISTATE_ENABLED) )
		clr = pTabSkin->m_colDisable;
	else if (m_fsState & TCISTATE_INACTIVE)
		clr = bAnchorColor ? pTabSkin->m_colInActive	: pTabSkin->m_colText;
	else
		clr = bAnchorColor ? pTabSkin->m_colActive		: pTabSkin->m_colText;

	COLORREF	clrOld = dc.SetTextColor(clr);

	CRect		rcBtn(m_rcItem.left + ptOffset.x, m_rcItem.top + ptOffset.y, m_rcItem.right, m_rcItem.bottom);
	rcBtn.DeflateRect(2, 0);

	UINT	uFormat;
	int		nWidth	= MtlComputeWidthOfText(m_strItem, dc.GetCurrentFont());

	if ( 1 /*nWidth > rcBtn.Width()*/ ) {
		uFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_LEFT | DT_END_ELLIPSIS;
	} else {
		uFormat = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_CENTER | DT_NOCLIP;
	}

	if ( !(m_fsState & TCISTATE_ENABLED) ) {
		CRect		rcDisabled	= rcBtn + CPoint(1, 1);
		COLORREF	clrOld2		= dc.SetTextColor(pTabSkin->m_colDisableHi);
		dc.DrawText(m_strItem, -1, rcDisabled, uFormat);
		dc.SetTextColor(clrOld2);
	}

	// �^�u��ɕ����������
	dc.DrawText(m_strItem, -1, rcBtn, uFormat);
	dc.SetTextColor(clrOld);
}



void	CTabCtrlItem::DrawHot(CDCHandle dc)
{
	dc.DrawEdge(m_rcItem, BDR_RAISEDINNER, BF_RECT);
}

void	CTabCtrlItem::DrawPressed(CDCHandle dc)
{
	dc.DrawEdge(m_rcItem, EDGE_SUNKEN, BF_MIDDLE);
}

void	CTabCtrlItem::DrawSelected(CDCHandle dc)
{
	dc.DrawEdge(m_rcItem, EDGE_SUNKEN, BF_RECT);
}

#endif













////////////////////////////////
// MtlCom����

// Helper for creating default FORMATETC from cfFormat
LPFORMATETC _MtlFillFormatEtc(LPFORMATETC lpFormatEtc, CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtcFill)
{
	ATLASSERT(lpFormatEtcFill != NULL);

	if (lpFormatEtc == NULL && cfFormat != 0) {
		lpFormatEtc 		  = lpFormatEtcFill;
		lpFormatEtc->cfFormat = cfFormat;
		lpFormatEtc->ptd	  = NULL;
		lpFormatEtc->dwAspect = DVASPECT_CONTENT;
		lpFormatEtc->lindex   = -1;
		lpFormatEtc->tymed	  = (DWORD) -1;
	}

	return lpFormatEtc;
}

bool MtlIsDataAvailable(IDataObject *pDataObject, CLIPFORMAT cfFormat, LPFORMATETC lpFormatEtc)
{
	ATLASSERT(pDataObject != NULL);

	// fill in FORMATETC struct
	FORMATETC	formatEtc;
	lpFormatEtc = MTL::_MtlFillFormatEtc(lpFormatEtc, cfFormat, &formatEtc);

	// attempt to get the data
	return pDataObject->QueryGetData(lpFormatEtc) == S_OK;
}

// MtlCom
////////////////////////////////




}	// namespace MTL









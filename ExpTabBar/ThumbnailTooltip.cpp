/**
*	@file	ThumbnailTooltip.cpp
*	@brief	サムネイルを表示するツールチップウィンドウ
*/

#include "stdafx.h"
#include "ThumbnailTooltip.h"
#include "GdiplusUtil.h"
#include "ShellWrap.h"
#include "ExpTabBarOption.h"

enum {
	//kMaxImageCache = 64,
	//kMaxImageWidth = 512,
	//kMaxImageHeight = 256,
	kBoundMargin = 4,
	kImageTextMargin = 8,
	kLeftTextMargin = kBoundMargin + 2,
	kTooltipTopItemMargin = 4,
};

CThumbnailTooltip::CThumbnailTooltip()
	: m_pNowImageData(nullptr)
{
	SetThemeClassList(VSCLASS_TOOLTIP);
}

CThumbnailTooltip::~CThumbnailTooltip()
{
	for (auto it = m_mapImageCache.begin(); it != m_mapImageCache.end(); ++it) {
		delete it->second.thumbnail;
	}
}

bool	CThumbnailTooltip::ShowThumbnailTooltip(std::wstring path, CRect rcItem)
{
	if (CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged) {
		// 作成するサムネイルの大きさが変わったのでキャッシュをクリアする
		m_mapImageCache.clear();
		CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged = false;
	}

	auto it = m_mapImageCache.find(path);
	if (it != m_mapImageCache.end()) {
		m_pNowImageData = &it->second;
	} else {
		ImageData	imgdata;
		std::unique_ptr<Gdiplus::Image> bmpRaw(Gdiplus::Bitmap::FromFile(path.c_str()));
		if (bmpRaw) {
			imgdata.ActualSize = _CalcActualSize(bmpRaw.get());

			/* サムネイル作成 */
			Gdiplus::Graphics	graphics(m_hWnd);
			imgdata.thumbnail = new Gdiplus::Bitmap(imgdata.ActualSize.cx, imgdata.ActualSize.cy, &graphics);
			Gdiplus::Graphics	graphicsTarget(imgdata.thumbnail);
			graphicsTarget.SetInterpolationMode(Gdiplus::InterpolationModeBilinear);
			graphicsTarget.DrawImage(bmpRaw.get(), 0, 0, imgdata.ActualSize.cx, imgdata.ActualSize.cy);

			//imgdata.thumbnail = bmpRaw->GetThumbnailImage(imgdata.ActualSize.cx, imgdata.ActualSize.cy);
			imgdata.strInfoTipText = ShellWrap::GetInfoTipText(path.c_str());
			imgdata.nInfoTipHeight = _CalcInfoTipTextHeight(imgdata);
	
			if (m_mapImageCache.size() > CThumbnailTooltipConfig::s_nMaxThumbnailCache) {
				delete m_mapImageCache.begin()->second.thumbnail;
				m_mapImageCache.erase(m_mapImageCache.begin());
			}
			auto it = m_mapImageCache.insert(std::make_pair<std::wstring, ImageData>(path, imgdata));
			m_pNowImageData = &it.first->second;
		}
	}
	if (m_pNowImageData) {
		CRect rcTooltip = _CalcTooltipRect(rcItem, *m_pNowImageData);
		MoveWindow(&rcTooltip);

		Invalidate(FALSE);
		SetWindowPos(HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_NOCOPYBITS);
		ShowWindow(SW_SHOWNOACTIVATE);
		return true;
	}
	return false;
}

void	CThumbnailTooltip::HideThumbnailTooltip()
{
	if (IsWindowVisible() == false)
		return ;

	ShowWindow(FALSE);

	m_pNowImageData = nullptr;
}


void CThumbnailTooltip::DoPaint(CDCHandle dc)
{
	if (m_pNowImageData == nullptr)
		return ;

	RECT rcClient;
	GetClientRect(&rcClient);
	--rcClient.right;
	--rcClient.bottom;
	CRect rcText = rcClient;
	rcText.top	= kBoundMargin + m_pNowImageData->ActualSize.cy + kImageTextMargin;
	rcText.left	= kLeftTextMargin;
	rcText.right -= kBoundMargin;
	rcText.bottom-= kBoundMargin;

	if (IsThemeNull() == false) {
		DrawThemeBackground(dc, TTP_STANDARD, TTCS_NORMAL, &rcClient);

		DrawThemeText(dc, TTP_STANDARD, TTCS_NORMAL, m_pNowImageData->strInfoTipText, m_pNowImageData->strInfoTipText.GetLength(), DT_END_ELLIPSIS, 0, &rcText);
	} else {
		dc.FillRect(&rcClient, COLOR_INFOBK);
		dc.DrawText(m_pNowImageData->strInfoTipText, m_pNowImageData->strInfoTipText.GetLength(), &rcText, DT_END_ELLIPSIS);
	}

	if (m_pNowImageData->thumbnail) {
		using namespace Gdiplus;
		Graphics	graphics(dc);
		graphics.DrawImage(m_pNowImageData->thumbnail, (REAL)kBoundMargin,(REAL)kBoundMargin/*,
			(REAL)m_pNowImageData->ActualSize.cx,(REAL)m_pNowImageData->ActualSize.cy*/);
	}
}


int CThumbnailTooltip::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	CLogFont	lf;
	lf.SetCaptionFont();
	m_NoThemeFont = lf.CreateFontIndirect();

	CThumbnailTooltipConfig::LoadConfig();

	return 0;
}

/// windows7のツールチップのように丸みをつける
void CThumbnailTooltip::OnSize(UINT nType, CSize size)
{
	SetWindowRgn(CreateRoundRectRgn( 0, 0, size.cx, size.cy, 3, 3 ), true);
}


/// モニターからはみ出ないようにツールチップの表示位置を計算する
CRect	CThumbnailTooltip::_CalcTooltipRect(const CRect& rcItem, const ImageData& ImageData)
{
	HMONITOR hMoni = MonitorFromRect(&rcItem, MONITOR_DEFAULTTONEAREST);
	MONITORINFO moniInfo = { sizeof(MONITORINFO) };
	GetMonitorInfo(hMoni, &moniInfo);

	int nToolTipWidth  = ImageData.ActualSize.cx + kBoundMargin*2 + 1;
	int nToolTipHeight = ImageData.ActualSize.cy + kBoundMargin*2 + 1 + kImageTextMargin + ImageData.nInfoTipHeight;
	CRect rcTooltip(rcItem.BottomRight(), CSize(nToolTipWidth, nToolTipHeight));

	CRect rcWork = moniInfo.rcWork;
	// モニターの右端をはみ出てる
	if (rcWork.right < rcTooltip.right && rcWork.bottom < rcTooltip.bottom) {
		rcTooltip.MoveToXY(rcItem.right - nToolTipWidth, rcItem.top - nToolTipHeight - kTooltipTopItemMargin);

	} else {
		if (rcWork.right < rcTooltip.right) {	// モニター右をはみ出る
			rcTooltip.MoveToX(rcWork.right - nToolTipWidth);
		}
		if (rcWork.bottom < rcTooltip.bottom) {	// モニター下をはみ出る
			rcTooltip.MoveToY(rcWork.bottom - nToolTipHeight);
		}
	}

	return rcTooltip;
}

/// 最大サイズに収まるように画像の比率を考えて縮小する
CSize	CThumbnailTooltip::_CalcActualSize(Gdiplus::Image* image)
{
	CSize ActualSize;
	int nImageWidth = image->GetWidth();
	int nImageHeight = image->GetHeight();
	const int kMaxImageWidth = CThumbnailTooltipConfig::s_MaxThumbnailSize.cx;
	const int kMaxImageHeight = CThumbnailTooltipConfig::s_MaxThumbnailSize.cy;
	if (nImageWidth > kMaxImageWidth || nImageHeight > kMaxImageHeight) {
		if (nImageHeight > kMaxImageHeight) {
			ActualSize.cx = (int)( (kMaxImageHeight / (double)nImageHeight) * nImageWidth );
			ActualSize.cy = kMaxImageHeight;
			if (ActualSize.cx > kMaxImageWidth) {
				ActualSize.cy = (int)( (kMaxImageWidth / (double)ActualSize.cx) * ActualSize.cy);
				ActualSize.cx = kMaxImageWidth;
			}
		} else {
			ActualSize.cy = (int)( (kMaxImageWidth / (double)nImageWidth) * nImageHeight );
			ActualSize.cy = kMaxImageWidth;
		}
	} else {
		ActualSize.SetSize(nImageWidth, nImageHeight);
	}
	return ActualSize;
}


int		CThumbnailTooltip::_CalcInfoTipTextHeight(const ImageData& ImageData)
{
	CDC dc = GetDC();
	if (IsThemeNull() == false) {
		CRect rcText;
		GetThemeTextExtent(dc, TTP_STANDARD, TTCS_NORMAL, 
			ImageData.strInfoTipText, 
			ImageData.strInfoTipText.GetLength(), 
			DT_END_ELLIPSIS, NULL, &rcText);
		return rcText.Height();
	} else {
		HFONT hFontPrev = dc.SelectFont(m_NoThemeFont);
		CSize	size;
		dc.GetTextExtent(ImageData.strInfoTipText, ImageData.strInfoTipText.GetLength(), &size);
		dc.SelectFont(hFontPrev);
		return size.cy;
	}
	
}



/**
*	@file	ThumbnailTooltip.cpp
*	@brief	サムネイルを表示するツールチップウィンドウ
*/

#include "stdafx.h"
#include "ThumbnailTooltip.h"
#include "GdiplusUtil.h"
#include "ShellWrap.h"
#include "ExpTabBarOption.h"
#include "Misc.h"
#include "Logger.h"

using namespace Gdiplus;

enum {
	//kMaxImageCache = 64,
	//kMaxImageWidth = 512,
	//kMaxImageHeight = 256,
	kBoundMargin = 4,
	kImageTextMargin = 8,
	kLeftTextMargin = kBoundMargin + 2,
	kTooltipTopItemMargin = 4,
};


//////////////////////////////////////////////////////////////////////
// CThumbnailTooltip

CThumbnailTooltip::CThumbnailTooltip()
	: m_pNowImageData(nullptr), m_nFramePosition(0), m_TimerID(0), m_bAddImageCached(false)
{
	SetThemeClassList(VSCLASS_TOOLTIP);

	m_bAlive = true;
}

CThumbnailTooltip::~CThumbnailTooltip()
{
}

bool	CThumbnailTooltip::ShowThumbnailTooltip(std::wstring path, CRect rcItem)
{
	if (CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged) {
		// 作成するサムネイルの大きさが変わったのでキャッシュをクリアする
		_ClearImageCache();
		CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged = false;
	}

	m_nFramePosition = 0;
	if (m_TimerID) {
		KillTimer(m_TimerID);
		m_TimerID = 0;
	}

	// Altキーを押しているときだけサムネイルを表示する
	if (CThumbnailTooltipConfig::s_bShowThumbnailOnAlt && (GetKeyState(VK_MENU) < 0) == false) {
		return false;
	}

	m_currentThumbnailPath = path;

	CCritSecLock	lock(m_cs);
	auto it = m_ImageCache.find(path);
	if (it != m_ImageCache.end()) {
		m_pNowImageData = it->get();
	} else {
		//ShowWindow(FALSE);
		m_pNowImageData = nullptr;
		Invalidate(FALSE);

		lock.Unlock();
		AddThumbnailCache(path, rcItem);
		return true;
	}
	if (m_pNowImageData) {
		CRect rcTooltip = _CalcTooltipRect(rcItem, *m_pNowImageData);
		//SetLayeredWindowAttributes(m_hWnd, 0, 0, LWA_ALPHA);
		MoveWindow(&rcTooltip, FALSE);

		if (m_pNowImageData->bGifAnimation) {
			m_TimerID = SetTimer(1, m_pNowImageData->vecDelayTime[0]);
		}

		/// windows7のツールチップのように丸みをつける
		SetWindowRgn(CreateRoundRectRgn(0, 0, rcTooltip.Width(), rcTooltip.Height(), 3, 3));

		SetWindowPos(HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE | SWP_NOCOPYBITS);
		ShowWindow(SW_SHOWNOACTIVATE);
		Invalidate(FALSE);
		UpdateWindow();
		//SetLayeredWindowAttributes(m_hWnd, 0, 255, LWA_ALPHA);
		return true;
	}
	return false;
}

// キャッシュの作成が完了したので、表示を行う
LRESULT CThumbnailTooltip::OnShowThumbnailWindowFromThread(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	const std::wstring& path = *(std::wstring*)wParam;
	auto itfound = m_CreateImageData.find(path);
	if (itfound != m_CreateImageData.end()) {
		if (path == m_currentThumbnailPath) {
			ShowThumbnailTooltip(path, itfound->get()->rcItem);
		}
		m_threadCreateImageData.detach();
		m_CreateImageData.erase(itfound);
		_StartCreateImageDataThread();
	} else {
		ATLASSERT(FALSE);
	}
	return 0;
}

void	CThumbnailTooltip::HideThumbnailTooltip()
{
	m_currentThumbnailPath.clear();

	if ( IsWindowVisible() == false )
		return ;

	ShowWindow(FALSE);

	if (m_TimerID) {
		KillTimer(m_TimerID);
		m_TimerID = 0;
	}

	m_pNowImageData = nullptr;
}

void	CThumbnailTooltip::_StartCreateImageDataThread()
{
	if (m_CreateImageData.size() == 0 || m_threadCreateImageData.joinable())
		return;

	auto& seqList = m_CreateImageData.get<seq>();
	auto path = seqList.front()->path;
	auto rcItem = seqList.front()->rcItem;
	m_threadCreateImageData = std::thread([this, path, rcItem]() {
		::CoInitialize(NULL);
		auto start = TimerStart();
		if (std::unique_ptr<ImageData> pdata = _CreateImageData(path.c_str())) {
			//INFO_LOG << L"[" << ::PathFindFileName(path.c_str()) << L"] _CreateImageData: " << ProcessingTimeStr(start);
			if (m_bAlive) {
				CCritSecLock	lock(m_cs);
				// キャッシュのサイズを超えたので削除
				if (m_bAddImageCached == false &&
					m_ImageCache.size() > CThumbnailTooltipConfig::s_nMaxThumbnailCache)
				{
					auto& seqList = m_ImageCache.get<seq>();
					seqList.erase(seqList.begin());
				}
				pdata->path = path;
				auto it = m_ImageCache.emplace(std::move(pdata));
			}
		} else {
			ATLASSERT(FALSE);
		}
		if (m_bAlive)
			SendMessage(WM_SHOWTHUMBNAILWINDOWFROMTHREAD, (WPARAM)&path);
		::CoUninitialize();
	});
}

void	CThumbnailTooltip::AddThumbnailCache(std::wstring path, CRect rcItem /*= CRect()*/)
{
	CCritSecLock	lock(m_cs);
	auto it = m_ImageCache.find(path);
	if (it == m_ImageCache.end()) {

		// 二重に作成しないようにする
		auto itcreateImage = m_CreateImageData.find(path);
		if (itcreateImage != m_CreateImageData.end()) {
			if (rcItem != CRect()) {
				m_CreateImageData.erase(itcreateImage);

				// 先頭に入れる
				auto& seqList = m_CreateImageData.get<seq>();
				seqList.emplace_front(new CreateImageData(path, rcItem));
				//INFO_LOG << L"キャッシュ作成の先頭に入れました[" << ::PathFindFileName(path.c_str()) << L"]";
			}
			return;
		}

		if (rcItem != CRect()) {
			// 先頭に入れる
			auto& seqList = m_CreateImageData.get<seq>();
			seqList.emplace_front(new CreateImageData(path, rcItem));
			//INFO_LOG << L"キャッシュ作成の先頭に入れました[" << ::PathFindFileName(path.c_str()) << L"]";
		} else {
			m_CreateImageData.emplace(new CreateImageData(path, rcItem));
		}
		_StartCreateImageDataThread();
	}
}

void	CThumbnailTooltip::OnLocationChanged()
{
	if (m_bAddImageCached) {
		_ClearImageCache();
		m_bAddImageCached = false;
	}
}


void CThumbnailTooltip::DoPaint(CDCHandle dc, RECT& /*rect*/)
{
	if (m_pNowImageData == nullptr) {
		RECT rcClient;
		GetClientRect(&rcClient);
		--rcClient.right;
		--rcClient.bottom;

		CRect rcText = rcClient;
		rcText.top = kBoundMargin;
		rcText.left = kLeftTextMargin;
		rcText.right -= kBoundMargin;
		rcText.bottom -= kBoundMargin;

		CString loadingText;
		loadingText.Format(_T("読み込み中...\r\n%s\r\n\r\n cachingFile (%d)"),
			::PathFindFileName(m_currentThumbnailPath.c_str()), m_CreateImageData.size());

		if (IsThemeNull() == false) {
			DrawThemeBackground(dc, TTP_STANDARD, TTCS_NORMAL, &rcClient);

			DrawThemeText(dc, TTP_STANDARD, TTCS_NORMAL, loadingText, loadingText.GetLength(), DT_WORDBREAK, 0, &rcText);
		} else {
			dc.FillRect(&rcClient, COLOR_INFOBK);
			dc.DrawText(loadingText, loadingText.GetLength(), &rcText, DT_WORDBREAK);
		}
		return;
	}

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

	Graphics	graphics(dc);
	if (m_pNowImageData->thumbnail) {
		graphics.DrawImage(m_pNowImageData->thumbnail, (REAL)kBoundMargin, (REAL)kBoundMargin);
	} else if (m_pNowImageData->bGifAnimation) {
		graphics.DrawImage(m_pNowImageData->vecGifImage[m_nFramePosition], (REAL)kBoundMargin, (REAL)kBoundMargin);
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

void CThumbnailTooltip::OnDestroy()
{
	m_bAlive = false;
	_ClearImageCache();
}

/// windows7のツールチップのように丸みをつける
void CThumbnailTooltip::OnSize(UINT nType, CSize size)
{
	
}

void CThumbnailTooltip::OnTimer(UINT_PTR nIDEvent)
{
	if (nIDEvent == 1 && m_pNowImageData) {
		++m_nFramePosition;
		if (m_nFramePosition == m_pNowImageData->nFrameCount)
			m_nFramePosition = 0;
		m_TimerID = SetTimer(1, m_pNowImageData->vecDelayTime[m_nFramePosition]);
		Invalidate(FALSE);
	}
}


/// モニターからはみ出ないようにツールチップの表示位置を計算する
CRect	CThumbnailTooltip::_CalcTooltipRect(const CRect& rcItem, const ImageData& ImageData)
{
	HMONITOR hMoni = MonitorFromRect(&rcItem, MONITOR_DEFAULTTONEAREST);
	MONITORINFO moniInfo = { sizeof(MONITORINFO) };
	GetMonitorInfo(hMoni, &moniInfo);

	int nToolTipWidth  = std::max(ImageData.ActualSize.cx, ImageData.InfoTipTextSize.cx) + kBoundMargin*2 + 1;
	int nToolTipHeight = ImageData.ActualSize.cy + kBoundMargin*2 + 1 + kImageTextMargin + ImageData.InfoTipTextSize.cy;
	CRect rcTooltip(rcItem.BottomRight(), CSize(nToolTipWidth, nToolTipHeight));

	bool bNoIcon = m_funcIsFolderDetailView();

	CRect rcWork = moniInfo.rcWork;
	// モニターの右端と下端をはみ出てる
	if (rcWork.right < rcTooltip.right && rcWork.bottom < rcTooltip.bottom) {
		if (bNoIcon) {
			// 右端の位置をアイテムの幅の3/4分左に移す
			rcTooltip.MoveToX(rcTooltip.left - static_cast<int>(rcItem.Width() * (3.0 / 4.0)));
			if (rcWork.right < rcTooltip.right) {
				rcTooltip.MoveToXY(rcItem.left - nToolTipWidth, rcWork.bottom - nToolTipHeight);
			} else {
				rcTooltip.MoveToXY(rcWork.right - nToolTipWidth, rcWork.bottom - nToolTipHeight);
			}
		} else {
			rcTooltip.MoveToXY(rcItem.left - nToolTipWidth, rcWork.bottom - nToolTipHeight);
		}
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
			ActualSize.cx = kMaxImageWidth;
		}
	} else {
		ActualSize.SetSize(nImageWidth, nImageHeight);
	}
	return ActualSize;
}


CSize		CThumbnailTooltip::_CalcInfoTipTextSize(const ImageData& ImageData)
{
	CDC dc = GetDC();
	if (IsThemeNull() == false) {
		CRect rcText;
		GetThemeTextExtent(dc, TTP_STANDARD, TTCS_NORMAL, 
			ImageData.strInfoTipText, 
			ImageData.strInfoTipText.GetLength(), 
			DT_END_ELLIPSIS, NULL, &rcText);
		return rcText.Size();
	} else {
		HFONT hFontPrev = dc.SelectFont(m_NoThemeFont);
		CSize	size;
		dc.GetTextExtent(ImageData.strInfoTipText, ImageData.strInfoTipText.GetLength(), &size);
		dc.SelectFont(hFontPrev);
		return size;
	}
	
}

/// 明示的に strPath のイメージキャッシュを作成する
std::unique_ptr<CThumbnailTooltip::ImageData>	CThumbnailTooltip::_CreateImageData(LPCTSTR strPath)
{
	std::unique_ptr<ImageData>	pdata(new ImageData);
	ImageData&	imgdata = *pdata;
	std::unique_ptr<Gdiplus::Image> bmpRaw(Gdiplus::Bitmap::FromFile(strPath));
	if (bmpRaw) {
		imgdata.ActualSize = _CalcActualSize(bmpRaw.get());

		/* サムネイル作成 */
		Gdiplus::Graphics	graphics(m_hWnd);
		if (Misc::GetPathExtention(strPath).CompareNoCase(_T("gif")) == 0) {
			GUID	guid;
			bmpRaw->GetFrameDimensionsList(&guid, 1);
			UINT	FrameCount = bmpRaw->GetFrameCount(&guid);
			if (FrameCount > 1) {
				imgdata.bGifAnimation = true;
				imgdata.nFrameCount = FrameCount;
				imgdata.vecGifImage.reserve(FrameCount);

				UINT	propItemSize = bmpRaw->GetPropertyItemSize(PropertyTagFrameDelay);
				PropertyItem*	propItems = (PropertyItem*)new BYTE[propItemSize];
				bmpRaw->GetPropertyItem(PropertyTagFrameDelay, propItemSize, propItems);
				for (UINT i = 0; i < FrameCount; ++i) {
					int nDelay = ((long*)propItems->value)[i] * 10;
					if (nDelay == 0)
						nDelay = 100;
					imgdata.vecDelayTime.push_back(nDelay);

					bmpRaw->SelectActiveFrame(&guid, i);
					Gdiplus::Image* imgPage = new Gdiplus::Bitmap(imgdata.ActualSize.cx, imgdata.ActualSize.cy, &graphics);
					Gdiplus::Graphics	graphicsTarget(imgPage);
					graphicsTarget.SetInterpolationMode(Gdiplus::InterpolationModeBilinear);
					graphicsTarget.DrawImage(bmpRaw.get(), 0, 0, imgdata.ActualSize.cx, imgdata.ActualSize.cy);
					imgdata.vecGifImage.push_back(imgPage);
				}
				delete[] (BYTE*)propItems;
				//imgdata.thumbnail	= bmpRaw->Clone();
			}
		}
		if (imgdata.bGifAnimation == false) {
			imgdata.thumbnail = new Gdiplus::Bitmap(imgdata.ActualSize.cx, imgdata.ActualSize.cy, &graphics);
			Gdiplus::Graphics	graphicsTarget(imgdata.thumbnail);
			graphicsTarget.SetInterpolationMode(Gdiplus::InterpolationModeBilinear);
			graphicsTarget.DrawImage(bmpRaw.get(), 0, 0, imgdata.ActualSize.cx, imgdata.ActualSize.cy);
		}

		//imgdata.thumbnail = bmpRaw->GetThumbnailImage(imgdata.ActualSize.cx, imgdata.ActualSize.cy);
		imgdata.strInfoTipText = ShellWrap::GetInfoTipText(strPath);
		imgdata.InfoTipTextSize = _CalcInfoTipTextSize(imgdata);
	}
	return pdata;
}


/// イメージキャッシュをクリアする
void	CThumbnailTooltip::_ClearImageCache()
{
	if (m_threadCreateImageData.joinable())
		m_threadCreateImageData.detach();
	m_CreateImageData.clear();

	{
		CCritSecLock	lock(m_cs);
		m_ImageCache.clear();
	}

	m_pNowImageData = nullptr;
	m_currentThumbnailPath.clear();
}


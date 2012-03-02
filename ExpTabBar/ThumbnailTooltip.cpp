/**
*	@file	ThumbnailTooltip.cpp
*	@brief	�T���l�C����\������c�[���`�b�v�E�B���h�E
*/

#include "stdafx.h"
#include "ThumbnailTooltip.h"
#include "GdiplusUtil.h"
#include "ShellWrap.h"
#include "ExpTabBarOption.h"
#include "Misc.h"

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

CThumbnailTooltip::CThumbnailTooltip()
	: m_pNowImageData(nullptr), m_nFramePosition(0), m_TimerID(0)
{
	SetThemeClassList(VSCLASS_TOOLTIP);
}

CThumbnailTooltip::~CThumbnailTooltip()
{
	_ClearImageCache();
}

bool	CThumbnailTooltip::ShowThumbnailTooltip(std::wstring path, CRect rcItem)
{
	if (CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged) {
		// �쐬����T���l�C���̑傫�����ς�����̂ŃL���b�V�����N���A����
		_ClearImageCache();
		CThumbnailTooltipConfig::s_bMaxThumbnailSizeChanged = false;
	}

	m_nFramePosition = 0;
	if (m_TimerID) {
		KillTimer(m_TimerID);
		m_TimerID = 0;
	}

	auto it = m_mapImageCache.find(path);
	if (it != m_mapImageCache.end()) {
		m_pNowImageData = &it->second;
	} else {
		ImageData	imgdata;
		std::unique_ptr<Gdiplus::Image> bmpRaw(Gdiplus::Bitmap::FromFile(path.c_str()));
		if (bmpRaw) {
			imgdata.ActualSize = _CalcActualSize(bmpRaw.get());

			/* �T���l�C���쐬 */
			Gdiplus::Graphics	graphics(m_hWnd);
			if (Misc::GetPathExtention(path.c_str()).CompareNoCase(_T("gif")) == 0) {
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
					delete propItems;
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

		if (m_pNowImageData->bGifAnimation) {
			m_TimerID = SetTimer(1, m_pNowImageData->vecDelayTime[0]);
		}

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

	if (m_TimerID) {
		KillTimer(m_TimerID);
		m_TimerID = 0;
	}

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

/// windows7�̃c�[���`�b�v�̂悤�Ɋۂ݂�����
void CThumbnailTooltip::OnSize(UINT nType, CSize size)
{
	SetWindowRgn(CreateRoundRectRgn( 0, 0, size.cx, size.cy, 3, 3 ), true);
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


/// ���j�^�[����͂ݏo�Ȃ��悤�Ƀc�[���`�b�v�̕\���ʒu���v�Z����
CRect	CThumbnailTooltip::_CalcTooltipRect(const CRect& rcItem, const ImageData& ImageData)
{
	HMONITOR hMoni = MonitorFromRect(&rcItem, MONITOR_DEFAULTTONEAREST);
	MONITORINFO moniInfo = { sizeof(MONITORINFO) };
	GetMonitorInfo(hMoni, &moniInfo);

	int nToolTipWidth  = ImageData.ActualSize.cx + kBoundMargin*2 + 1;
	int nToolTipHeight = ImageData.ActualSize.cy + kBoundMargin*2 + 1 + kImageTextMargin + ImageData.nInfoTipHeight;
	CRect rcTooltip(rcItem.BottomRight(), CSize(nToolTipWidth, nToolTipHeight));

	CRect rcWork = moniInfo.rcWork;
	// ���j�^�[�̉E�[���͂ݏo�Ă�
	if (rcWork.right < rcTooltip.right && rcWork.bottom < rcTooltip.bottom) {
		rcTooltip.MoveToXY(rcItem.right - nToolTipWidth, rcItem.top - nToolTipHeight - kTooltipTopItemMargin);

	} else {
		if (rcWork.right < rcTooltip.right) {	// ���j�^�[�E���͂ݏo��
			rcTooltip.MoveToX(rcWork.right - nToolTipWidth);
		}
		if (rcWork.bottom < rcTooltip.bottom) {	// ���j�^�[�����͂ݏo��
			rcTooltip.MoveToY(rcWork.bottom - nToolTipHeight);
		}
	}

	return rcTooltip;
}

/// �ő�T�C�Y�Ɏ��܂�悤�ɉ摜�̔䗦���l���ďk������
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

/// �C���[�W�L���b�V�����N���A����
void	CThumbnailTooltip::_ClearImageCache()
{
	std::for_each(m_mapImageCache.begin(), m_mapImageCache.end(), [](std::pair<std::wstring, ImageData> pr) {
		ImageData& imgdata = pr.second;
		delete imgdata.thumbnail;
		if (imgdata.bGifAnimation) {
			std::for_each(imgdata.vecGifImage.begin(), imgdata.vecGifImage.end(), [](Gdiplus::Image* img) {
				delete img;
			});
		}
	});
	m_mapImageCache.clear();
}


/**
*	@file	ThumbnailTooltip.h
*	@brief	サムネイルを表示するツールチップウィンドウ
*/

#pragma once

#include <unordered_map>

namespace Gdiplus {
	class Image;
};

class CThumbnailTooltip : 
	public CDoubleBufferWindowImpl<CThumbnailTooltip>,
	public CThemeImpl<CThumbnailTooltip>
{
public:
	DECLARE_WND_CLASS_EX(NULL, CS_HREDRAW | CS_VREDRAW | CS_DROPSHADOW, 0);

	struct ImageData {
		Gdiplus::Image*	thumbnail;
		CSize	ActualSize;
		CString strInfoTipText;
		int		nInfoTipHeight;
		bool	bGifAnimation;
		GUID	dimentionID;
		std::vector<int>				vecDelayTime;
		UINT	nFrameCount;
		ImageData() : thumbnail(nullptr), nInfoTipHeight(0), bGifAnimation(false), nFrameCount(0) { }
	};

	CThumbnailTooltip();
	~CThumbnailTooltip();

	bool	ShowThumbnailTooltip(std::wstring path, CRect rcItem);
	void	HideThumbnailTooltip();

	// Overrides
	void DoPaint(CDCHandle dc);

	BEGIN_MSG_MAP( CThumbnailTooltip )
		CHAIN_MSG_MAP( CDoubleBufferWindowImpl<CThumbnailTooltip> )
		CHAIN_MSG_MAP( CThemeImpl<CThumbnailTooltip> )
		MSG_WM_CREATE( OnCreate )
		MSG_WM_SIZE	( OnSize )
		MSG_WM_TIMER( OnTimer )
	END_MSG_MAP()

	int OnCreate(LPCREATESTRUCT lpCreateStruct);
	void OnSize(UINT nType, CSize size);
	void OnTimer(UINT_PTR nIDEvent);

private:
	CRect	_CalcTooltipRect(const CRect& rcItem, const ImageData& ImageData);
	CSize	_CalcActualSize(Gdiplus::Image* image);
	int		_CalcInfoTipTextHeight(const ImageData& ImageData);

	// Data members
	CFont	m_NoThemeFont;
	ImageData*	m_pNowImageData;
	UINT		m_nFramePosition;
	UINT_PTR	m_TimerID;

	std::unordered_map<std::wstring, ImageData>	m_mapImageCache;
};



















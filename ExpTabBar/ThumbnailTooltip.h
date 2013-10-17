/**
*	@file	ThumbnailTooltip.h
*	@brief	サムネイルを表示するツールチップウィンドウ
*/

#pragma once

#include <unordered_map>
#include <memory>
#include <atlsync.h>

namespace Gdiplus {
	class Image;
};


class CThumbnailTooltip : 
	public CBufferedPaintWindowImpl<CThumbnailTooltip>,
	public CThemeImpl<CThumbnailTooltip>
{
public:
	DECLARE_WND_CLASS_EX(NULL, /*CS_HREDRAW | CS_VREDRAW |*/ CS_DROPSHADOW, 0);

	struct ImageData {
		Gdiplus::Image*	thumbnail;
		CSize	ActualSize;
		CString strInfoTipText;
		CSize	InfoTipTextSize;
		bool	bGifAnimation;
		std::vector<Gdiplus::Image*>	vecGifImage;
		std::vector<int>				vecDelayTime;
		UINT	nFrameCount;
		ImageData() : thumbnail(nullptr), bGifAnimation(false), nFrameCount(0) { }
	};

	CThumbnailTooltip();
	~CThumbnailTooltip();

	bool	ShowThumbnailTooltip(std::wstring path, CRect rcItem);
	void	HideThumbnailTooltip();

	void	AddThumbnailCache(LPCTSTR strPath);
	void	OnLocationChanged();

	// Overrides
	void DoPaint(CDCHandle dc, RECT& /*rect*/);

	BEGIN_MSG_MAP( CThumbnailTooltip )
		CHAIN_MSG_MAP(CBufferedPaintWindowImpl<CThumbnailTooltip>)
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
	CSize	_CalcInfoTipTextSize(const ImageData& ImageData);
	std::unique_ptr<ImageData>	_CreateImageData(LPCTSTR strPath);
	void	_ClearImageCache();

	// Data members
	CFont	m_NoThemeFont;
	ImageData*	m_pNowImageData;
	UINT		m_nFramePosition;
	UINT_PTR	m_TimerID;

	std::unordered_map<std::wstring, ImageData*>	m_mapImageCache;
	CCriticalSection	m_cs;
	bool	m_bAddImageCached;
};



















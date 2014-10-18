/**
*	@file	ThumbnailTooltip.h
*	@brief	サムネイルを表示するツールチップウィンドウ
*/

#pragma once

#include <unordered_map>
#include <memory>
#include <atlsync.h>
#include <thread>
#include <vector>
#include <boost\multi_index_container.hpp>
#include <boost\multi_index\sequenced_index.hpp>
#include <boost\multi_index\hashed_index.hpp>
#include <boost\multi_index\member.hpp>
#include "GdiplusUtil.h"

using namespace boost::multi_index;


class CThumbnailTooltip : 
	public CBufferedPaintWindowImpl<CThumbnailTooltip>,
	public CThemeImpl<CThumbnailTooltip>
{
public:
	DECLARE_WND_CLASS_EX(NULL, /*CS_HREDRAW | CS_VREDRAW |*/ CS_DROPSHADOW, 0);

	struct ImageData {
		std::wstring path;
		Gdiplus::Image*	thumbnail;
		CSize	ActualSize;
		CString strInfoTipText;
		CSize	InfoTipTextSize;
		bool	bGifAnimation;
		std::vector<Gdiplus::Image*>	vecGifImage;
		std::vector<int>				vecDelayTime;
		UINT	nFrameCount;

		ImageData() : thumbnail(nullptr), bGifAnimation(false), nFrameCount(0) { }

		~ImageData()
		{
			if (thumbnail)
				delete thumbnail;
			if (bGifAnimation) {
				for (auto img : vecGifImage)
					delete img;
			}
		}
	};

	struct CreateImageData {
		std::thread	createThread;
		std::wstring path;
		CRect rcItem;

		CreateImageData(std::thread&& thread, const std::wstring& path, const CRect& rcItem) 
			: createThread(std::move(thread)), path(path), rcItem(rcItem)
		{ }
	};

	
	struct order {}; // 辞書順のタグ
	struct seq {}; // 挿入順のタグ

	template<class T>
	using MapContainer = multi_index_container<
		std::unique_ptr<T>,
		indexed_by<
			hashed_unique<tag<order>, member<T, std::wstring, &T::path>>,
			sequenced<tag<seq> > // 挿入順
		>
	>;

	enum {
		WM_SHOWTHUMBNAILWINDOWFROMTHREAD = WM_APP + 100,
	};

	CThumbnailTooltip();
	~CThumbnailTooltip();

	bool	ShowThumbnailTooltip(std::wstring path, CRect rcItem);
	void	HideThumbnailTooltip();

	void	LockImageCache() { m_bAddImageCached = true; }
	void	AddThumbnailCache(std::wstring path, CRect rcItem = CRect());
	void	OnLocationChanged();
	void	ClearImageCache() { _ClearImageCache(); }

	// Overrides
	void DoPaint(CDCHandle dc, RECT& /*rect*/);

	BEGIN_MSG_MAP( CThumbnailTooltip )
		CHAIN_MSG_MAP(CBufferedPaintWindowImpl<CThumbnailTooltip>)
		CHAIN_MSG_MAP( CThemeImpl<CThumbnailTooltip> )
		MSG_WM_CREATE( OnCreate )
		MSG_WM_SIZE	( OnSize )
		MSG_WM_TIMER( OnTimer )
		MESSAGE_HANDLER_EX(WM_SHOWTHUMBNAILWINDOWFROMTHREAD, OnShowThumbnailWindowFromThread)
	END_MSG_MAP()

	int OnCreate(LPCREATESTRUCT lpCreateStruct);
	void OnSize(UINT nType, CSize size);
	void OnTimer(UINT_PTR nIDEvent);
	LRESULT OnShowThumbnailWindowFromThread(UINT uMsg, WPARAM wParam, LPARAM lParam);

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

	MapContainer<ImageData>	m_ImageCache;
	CCriticalSection	m_cs;
	bool	m_bAddImageCached;

	std::wstring	m_currentThumbnailPath;
	MapContainer<CreateImageData>	m_CreateImageData;
};



















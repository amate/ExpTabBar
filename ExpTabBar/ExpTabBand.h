// ExpTabBand.h : CExpTabBand の宣言

#pragma once
#include "ExpTabBar_i.h"
#include "resource.h"       // メイン シンボル
#include <comsvcs.h>
#include <exdispid.h>
#include <UIAutomation.h>
//#include "ReflectionWnd.h"
#include "DonutTabBar.h"
#include "ThumbnailTooltip.h"
#include "ExpTabBarOption.h"

#define SINKID_EVENTS	1234

////////////////////////////////////////////////////////////////////////////////
// CExpTabBand

class ATL_NO_VTABLE CExpTabBand :  
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CExpTabBand, &CLSID_ExpTabBand>,
	public IDeskBand,							// IEツールバーを作る際に必要
	public IObjectWithSiteImpl<CExpTabBand>,	// 
	public IDispatchImpl<IExpTabBand, &IID_IExpTabBand, &LIBID_ExpTabBarLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispEventImpl<SINKID_EVENTS, CExpTabBand, &DIID_DWebBrowserEvents2, &LIBID_SHDocVw, 1, 1>,
	public CWindowImpl<CExpTabBand>	// メッセージマップ用
{
public:
	// コンストラクタ/デストラクタ
	CExpTabBand();
	~CExpTabBand();

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct() { return S_OK; }
	void	FinalRelease() { }

DECLARE_REGISTRY_RESOURCEID(IDR_EXPTABBAND)

DECLARE_NOT_AGGREGATABLE(CExpTabBand)

BEGIN_COM_MAP(CExpTabBand)
	COM_INTERFACE_ENTRY(IExpTabBand)
    COM_INTERFACE_ENTRY(IOleWindow)
    COM_INTERFACE_ENTRY_IID(IID_IDockingWindow, IDockingWindow)
    COM_INTERFACE_ENTRY_IID(IID_IDeskBand, IDeskBand)
    COM_INTERFACE_ENTRY(IObjectWithSite)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


BEGIN_SINK_MAP(CExpTabBand)
    SINK_ENTRY_EX(SINKID_EVENTS, DIID_DWebBrowserEvents2, DISPID_TITLECHANGE		, OnTitleChange)
	SINK_ENTRY_EX(SINKID_EVENTS, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2	, OnNavigateComplete2)
	SINK_ENTRY_EX(SINKID_EVENTS, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE	, OnDocumentComplete)
END_SINK_MAP()


BEGIN_MSG_MAP(CExpTabBand)
ALT_MSG_MAP(1)
	NOTIFY_CODE_HANDLER_EX(LVN_ITEMCHANGED, OnListViewItemChanged)
	//NOTIFY_CODE_HANDLER_EX(TTN_SHOW, OnListViewGetDispInfo)
	MSG_WM_PARENTNOTIFY( OnParentNotify )
ALT_MSG_MAP(2)	// SysListView32
	if (CThumbnailTooltipConfig::s_bUseThumbnailTooltip) {
		NOTIFY_CODE_HANDLER_EX(TTN_GETDISPINFOW, OnListViewGetDispInfo)
		MSG_WM_MOUSEMOVE( OnListViewMouseMove )
		MSG_WM_MOUSELEAVE( OnListViewMouseLeave )
		MSG_WM_MOUSEHOVER( OnListViewMouseHover )
		MSG_WM_KEYUP	 ( OnListViewKeyUp	)
		MSG_WM_KILLFOCUS ( OnListViewKillFocus )
	}
ALT_MSG_MAP(3)	// DirectUI
	if (CThumbnailTooltipConfig::s_bUseThumbnailTooltip) {
		NOTIFY_CODE_HANDLER_EX(TTN_NEEDTEXTW, OnListViewGetDispInfo)
		MSG_WM_MOUSEMOVE( OnListViewMouseMove )
		MSG_WM_MOUSELEAVE( OnListViewMouseLeave )
		MSG_WM_MOUSEHOVER( OnListViewMouseHover )
		MSG_WM_KEYUP	 ( OnListViewKeyUp	)
		MSG_WM_KILLFOCUS ( OnListViewKillFocus )
	}
ALT_MSG_MAP(5)
	MSG_WM_LBUTTONDBLCLK( OnTabBarLButtonDblClk	)
END_MSG_MAP()

	BOOL IsMsgHandled() const { return FALSE; }	//
	void SetMsgHandled(_In_ BOOL bHandled) { }	// アサートが出るので何もしないようにする


// Interfaces
public:
	// IDeskBand
    STDMETHOD(GetBandInfo)(DWORD dwBandID, DWORD dwViewMode, DESKBANDINFO* pdbi);

	// IOleWindow
    STDMETHOD(GetWindow)(HWND* phwnd);
    STDMETHOD(ContextSensitiveHelp)(BOOL fEnterMode);

	// IDockingWindow
    STDMETHOD(ShowDW)(BOOL fShow);
    STDMETHOD(CloseDW)(unsigned long dwReserved);
    STDMETHOD(ResizeBorderDW)(const RECT* prcBorder, IUnknown* punkToolbarSite, BOOL fReserved);

	// IObjectWithSite methods
	STDMETHOD(SetSite) (IUnknown*);
	STDMETHOD(GetSite) (REFIID, LPVOID*);


	// Event sink
	void __stdcall OnNavigateComplete2(IDispatch* pDisp,VARIANT* URL);
	void __stdcall OnDocumentComplete(IDispatch* pDisp, VARIANT* URL);
    void __stdcall OnTitleChange(BSTR title);


	LRESULT OnListViewItemChanged(LPNMHDR pnmh);
	LRESULT OnListViewGetDispInfo(LPNMHDR pnmh);
	void	OnListViewMouseMove(UINT nFlags, CPoint point);
	void	OnListViewMouseLeave();
	void	OnListViewMouseHover(WPARAM wParam, CPoint ptPos);
	void	OnListViewKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags);
	void	OnListViewKillFocus(CWindow wndFocus);

	void	OnParentNotify(UINT message, UINT nChildID, LPARAM lParam);

	void	OnTabBarLButtonDblClk(UINT nFlags, CPoint point);

private:
	int		_HitTestDirectUI(CRect& rcItem);
	int		_HitTestListView(const CPoint& pt);
	int		_HitTestListView();
	CRect	_GetItemRect(int nIndex);
	bool	_ShowThumbnailTooltip(int nIndex, CRect rcItem);
	void	_HideThumbnailTooltip();
	void	_TrackMouseLeave(HWND hWnd);
	void	_TrackMouseHover(HWND hWnd);

	// Data members
	CDonutTabBar	m_wndTabBar;

	CComPtr<IShellBrowser>	m_spShellBrowser;
	CComPtr<ITravelLogStg>	m_spTravelLogStg;
	CComPtr<IWebBrowser2>	m_spWebBrowser2;
	CComPtr<IUIAutomation>	m_spUIAutomation;

	bool	m_bNavigateCompleted;
	CContainedWindow	m_wndShellView;
	CContainedWindow	m_wndListView;
	CContainedWindow	m_wndDirectUI;

	CListViewCtrl		m_ListView;
	CToolTipCtrl		m_Tooltip;
	int		m_nIndexTooltip;
	bool	m_bNowTrackMouseLeave;
	bool	m_bNowTrackMouseHover;
	CPoint	m_ptLastForMouseMove;

	CThumbnailTooltip	m_ThumbnailTooltip;
};

OBJECT_ENTRY_AUTO(__uuidof(ExpTabBand), CExpTabBand)





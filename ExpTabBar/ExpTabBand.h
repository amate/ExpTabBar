// ExpTabBand.h : CExpTabBand の宣言

#pragma once
#include "ExpTabBar_i.h"
#include "resource.h"       // メイン シンボル
#include <comsvcs.h>
#include <exdispid.h>

//#include "ReflectionWnd.h"
#include "DonutTabBar.h"



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
	MSG_WM_PARENTNOTIFY( OnParentNotify )
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
	void	OnParentNotify(UINT message, UINT nChildID, LPARAM lParam);

private:
	// Data members
	CDonutTabBar	m_wndTabBar;

	CComPtr<IShellBrowser>	m_spShellBrowser;
	CComPtr<ITravelLogStg>	m_spTravelLogStg;
	CComPtr<IWebBrowser2>	m_spWebBrowser2;

	bool	m_bNavigateCompleted;
	CContainedWindow	m_wndShellView;
	CListViewCtrl		m_ListView;
};

OBJECT_ENTRY_AUTO(__uuidof(ExpTabBand), CExpTabBand)





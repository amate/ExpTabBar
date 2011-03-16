// ExpTabBand.h : CExpTabBand の宣言

#pragma once
#include "ExpTabBar_i.h"
#include "resource.h"       // メイン シンボル
#include <comsvcs.h>

#include "ReflectionWnd.h"

#include <exdispid.h>

#define SINKID_EVENTS	1234

//////////////////////////////////////////////////////////////////////////
// CWebBrowserEvents : ブラウザで起こったイベントの発生を伝える

class CWebBrowserEvents :
	public CComObjectRootEx<CComSingleThreadModel>,
	public IDispEventImpl<SINKID_EVENTS, CWebBrowserEvents, &DIID_DWebBrowserEvents2, &LIBID_SHDocVw, 1, 1>
{
public:

	CDonutTabBar*	pTabBar;
	bool	m_bNavigateCompleted;

	// コンストラクタ
	CWebBrowserEvents()
		: m_bNavigateCompleted(false)
	{
	}

	void SetTabBarWindow(CDonutTabBar* p) { pTabBar = p; }

	BEGIN_COM_MAP(CWebBrowserEvents)
	  COM_INTERFACE_ENTRY_IID(DIID_DWebBrowserEvents2, CWebBrowserEvents)
	END_COM_MAP()

	BEGIN_SINK_MAP(CWebBrowserEvents)
        SINK_ENTRY_EX(SINKID_EVENTS, DIID_DWebBrowserEvents2, DISPID_TITLECHANGE, OnTitleChange)
		SINK_ENTRY_EX(SINKID_EVENTS, DIID_DWebBrowserEvents2, DISPID_NAVIGATECOMPLETE2, OnNavigateComplete2)
		SINK_ENTRY_EX(SINKID_EVENTS, DIID_DWebBrowserEvents2, DISPID_DOCUMENTCOMPLETE, OnDocumentComplete)
    END_SINK_MAP()

	void __stdcall OnNavigateComplete2(IDispatch *pDisp,VARIANT *URL)
	{
		m_bNavigateCompleted = true;
		
		CString strURL(*URL);
		ATLTRACE(_T(" URL : %s\n"), strURL);
		//MessageBox(NULL, strURL, NULL, NULL);
		pTabBar->NavigateComplete2(strURL);
	}

	void __stdcall OnDocumentComplete(IDispatch* pDisp, VARIANT* URL)
	{//MessageBox(NULL, _T("OnDocumentComplete"), NULL, NULL);
		if (m_bNavigateCompleted == true) {
			pTabBar->DocumentComplete();
			m_bNavigateCompleted = false;
		}
	}


    void __stdcall OnTitleChange(BSTR title)
	{//MessageBox(NULL, title, NULL, NULL);
		pTabBar->RefreshTab(title);
    }

};


////////////////////////////////////////////////////////////////////////////////
// CExpTabBand

class ATL_NO_VTABLE CExpTabBand :  
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CExpTabBand, &CLSID_ExpTabBand>,
	public IDeskBand,							// IEツールバーを作る際に必要
	public IObjectWithSiteImpl<CExpTabBand>,	// 
	public IDispatchImpl<IExpTabBand, &IID_IExpTabBand, &LIBID_ExpTabBarLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	HWND 					m_hWndParent;
	CReflectionWnd			m_wndReflection;
	CComPtr<IShellBrowser>	m_spShellBrowser;
	CComPtr<ITravelLogStg>	m_spTravelLogStg;
	CComPtr<IWebBrowser2>	m_spWebBrowser2;
	CComObject<CWebBrowserEvents>*	m_pSink;


public:
	// コンストラクタ/デストラクタ
	CExpTabBand();
	~CExpTabBand();

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

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
	STDMETHOD (SetSite) (IUnknown*);
	STDMETHOD (GetSite) (REFIID, LPVOID*);

// Implementation:
public:
	BOOL	RegisterAndCreateWindow();

};

OBJECT_ENTRY_AUTO(__uuidof(ExpTabBand), CExpTabBand)





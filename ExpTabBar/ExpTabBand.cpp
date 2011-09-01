// ExpTabBand.cpp : CExpTabBand の実装

#include "stdafx.h"
#include "ExpTabBand.h"
#include <UIAutomation.h>
#include <boost/lexical_cast.hpp>
#include "ShellWrap.h"



// コンストラクタ/デストラクタ
CExpTabBand::CExpTabBand() : m_bNavigateCompleted(false), m_wndShellView(this, 1)
{	}

CExpTabBand::~CExpTabBand()
{
	DispEventUnadvise(m_spWebBrowser2);
}


// IDeskBand
STDMETHODIMP CExpTabBand::GetBandInfo(DWORD /* dwBandID */, DWORD /* dwViewMode */, DESKBANDINFO* pdbi)
{
    if (pdbi) {
        //ツールバーの最小サイズ
        if (pdbi->dwMask & DBIM_MINSIZE) {
            pdbi->ptMinSize.x = -1;
            pdbi->ptMinSize.y = 24;
        }
        //ツールバーの最大サイズ
        if (pdbi->dwMask & DBIM_MAXSIZE) {
            pdbi->ptMaxSize.x = -1;
            pdbi->ptMaxSize.y = -1;
        }
        // ツールバーとビューとの間をドラッグしたときどれぐらいの間隔ずつ空けるか
		// dwModeFlags に DBIMF_VARIABLEHEIGHT が設定されていないと無視される
        if (pdbi->dwMask & DBIM_INTEGRAL) {
            pdbi->ptIntegral.x = -1;
            pdbi->ptIntegral.y = -1;
        }

		// 理想
        if (pdbi->dwMask & DBIM_ACTUAL) {
            pdbi->ptActual.x = -1;
            pdbi->ptActual.y = -1;
        }

		//ツールバー左側に表示されるタイトル文字列
        if (pdbi->dwMask & DBIM_TITLE) {
			//wcscpy_s(pdbi->wszTitle, _T("ExpTabBar"));
        }

        if( pdbi->dwMask & DBIM_BKCOLOR ) {
            //こう書いておけば、デフォルトの背景色になってくれるらしい
            pdbi->dwMask &= ~DBIM_BKCOLOR;
        }

        if (pdbi->dwMask & DBIM_MODEFLAGS) {
            pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_BREAK;	// バーを表示する際に、必ず次の行にしてくれる指定
        }
    }
    return S_OK;
}


// IOleWindow
STDMETHODIMP CExpTabBand::GetWindow(HWND* phwnd)
{
	if (NULL == phwnd) {
		return E_INVALIDARG;
	} else {
		*phwnd = m_wndTabBar.m_hWnd;
		return S_OK;
	}
}

STDMETHODIMP CExpTabBand::ContextSensitiveHelp(BOOL /* fEnterMode */)
{
    return E_NOTIMPL;
}


// IDockingWindow
STDMETHODIMP CExpTabBand::ShowDW(BOOL fShow)
{
    //ツールバーのShowWindowを実行する
    //まだツールバーを作ってないので、何もしない
	ATLTRACE(_T("ShowDW() : %s\n"), fShow ? _T("true") : _T("false"));
    return S_OK;
}

STDMETHODIMP CExpTabBand::CloseDW(unsigned long /* dwReserved */)
{
    //CLOSE時の動作。ツールバーを非表示にする
    ShowDW(FALSE);
	ATLTRACE(_T("CloseDW()\n"));
    return S_OK;
}

STDMETHODIMP CExpTabBand::ResizeBorderDW(const RECT* /* prcBorder */, IUnknown* /* punkToolbarSite */, BOOL /* fReserved */)
{
    return E_NOTIMPL;
}

///////////////////////////////////////////////////////////////////////////
//
// IObjectWithSite implementations
//

STDMETHODIMP CExpTabBand::SetSite(IUnknown* punkSite)
{
	HRESULT hr;
	if (punkSite) {
		//Get the parent window.
		CComQIPtr<IOleWindow> pOleWindow(punkSite);
		ATLASSERT(pOleWindow);

		HWND	hWndParent;
		pOleWindow->GetWindow(&hWndParent);
		if (hWndParent == NULL) {
			ATLASSERT(FALSE);
			return E_FAIL;
		}
		
		CComQIPtr<IServiceProvider> pSP(punkSite);
		ATLASSERT(pSP);
		hr = pSP->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (LPVOID*)&m_spWebBrowser2);
		ATLASSERT(m_spWebBrowser2);
		if (SUCCEEDED(hr)) {
			hr = DispEventAdvise(m_spWebBrowser2);
			if (FAILED(hr)) {
				ATLASSERT(FALSE);
				return E_FAIL;
			}
		}

		pSP->QueryService(SID_SShellBrowser, &m_spShellBrowser);
		ATLASSERT(m_spShellBrowser);
		if (m_spShellBrowser == nullptr) {
			ATLASSERT(FALSE);
			return E_FAIL;
		}


		m_wndTabBar.Initialize(punkSite);

		/* タブバー作成 */
		m_wndTabBar.Create(hWndParent);
		if (m_wndTabBar.IsWindow() == FALSE) {
			ATLASSERT(FALSE);
			return E_FAIL;
		}
	}

	return S_OK;
}

STDMETHODIMP CExpTabBand::GetSite(REFIID riid, LPVOID *ppvReturn)
{
	return E_FAIL;
}


// Event sink

void CExpTabBand::OnNavigateComplete2(IDispatch* pDisp,VARIANT* URL)
{
	m_bNavigateCompleted = true;
		
	CString strURL(*URL);
	ATLTRACE(_T(" URL : %s\n"), strURL);
	m_wndTabBar.NavigateComplete2(strURL);
}

void CExpTabBand::OnDocumentComplete(IDispatch* pDisp, VARIANT* URL)
{
	if (m_bNavigateCompleted == true) {
		m_wndTabBar.DocumentComplete();
		m_bNavigateCompleted = false;

		if (m_wndShellView.m_hWnd)
			m_wndShellView.UnsubclassWindow(TRUE);

		HWND hWnd;
		CComPtr<IShellView>	spShellView;
		HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
		if (SUCCEEDED(hr)) {
			hr = spShellView->GetWindow(&hWnd);
			if (SUCCEEDED(hr)) {
				m_wndShellView.SubclassWindow(hWnd);
				m_ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
			}
		}
	}
}


void CExpTabBand::OnTitleChange(BSTR title)
{
	m_wndTabBar.RefreshTab(title);
}


// Message map

/// 選択されているアイテムが変わった
LRESULT CExpTabBand::OnListViewItemChanged(LPNMHDR pnmh)
{
	m_wndShellView.DefWindowProc();

	LPNMLISTVIEW	pnmlv = (LPNMLISTVIEW)pnmh;
	if (pnmlv->iItem == -1)
		return 0;
	
	/* アイコン部分が再描写されないバグに対処 */
	RECT rcItem;
	m_ListView.GetItemRect(pnmlv->iItem, &rcItem, LVIR_ICON);
	m_ListView.InvalidateRect(&rcItem);
	ATLTRACE(_T("OnListViewItemChanged() : %d\n"), pnmlv->iItem);
	return 0;
}

/// フォルダーをミドルクリックで新しいタブで開く
void CExpTabBand::OnParentNotify(UINT message, UINT nChildID, LPARAM lParam)
{
	if (message == WM_MBUTTONDOWN) {
		int nIndex = -1;
		if (m_ListView.m_hWnd) {
			POINT pt;
			::GetCursorPos(&pt);
			m_ListView.ScreenToClient(&pt);
			UINT Flags;
			nIndex = m_ListView.HitTest(pt, &Flags);
		} else {
			HRESULT	hr;
			CComPtr<IUIAutomation>	spUIAutomation;
			hr = ::CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&spUIAutomation);
			ATLASSERT(spUIAutomation);

			CComPtr<IUIAutomationElement>	spShellViewElement;
			hr = spUIAutomation->ElementFromHandle(m_wndShellView, &spShellViewElement);
			if (FAILED(hr))
				return ;


			CComVariant	vProp(UIA_ListItemControlTypeId);
			CComPtr<IUIAutomationCondition>	spCondition;
			hr = spUIAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, vProp, &spCondition);
			if (FAILED(hr)) 
				return ;

			CComPtr<IUIAutomationElementArray>	spUIElementArray;
			hr = spShellViewElement->FindAll(TreeScope_Descendants, spCondition, &spUIElementArray);
			if (FAILED(hr) || spUIElementArray == nullptr)
				return ;

			POINT	pt;
			::GetCursorPos(&pt);
			
			int nCount = 0;
			spUIElementArray->get_Length(&nCount);
			for (int i = 0; i < nCount; ++i) {
				CComPtr<IUIAutomationElement>	spUIElm;
				spUIElementArray->GetElement(i, &spUIElm);
				BOOL	bOffScreen = FALSE;
				spUIElm->get_CurrentIsOffscreen(&bOffScreen);
				if (bOffScreen)
					continue;
				CRect rcItem;
				spUIElm->get_CurrentBoundingRectangle(&rcItem);
				if (rcItem.PtInRect(pt)) {	// 見つかった
					nIndex = i;
					break;
				}
			}

#if 0	//\\ こっちはちょっと問題がある
			POINT pt;
			::GetCursorPos(&pt);
			CComPtr<IUIAutomationElement>	spPointedElement;
			hr = spUIAutomation->ElementFromPoint(pt, &spPointedElement);
			if (FAILED(hr))
				return ;
			
			CComPtr<IUIAutomationElement>	spItemUIElement;	

			CComBSTR	strPointedClassName;
			spPointedElement->get_CurrentClassName(&strPointedClassName);
			if (strPointedClassName == nullptr || strPointedClassName != L"UIItem") {
				CComPtr<IUIAutomationTreeWalker>	spWalker;
				spUIAutomation->get_ControlViewWalker(&spWalker);
				if (spWalker == nullptr)
					return ;
				
				spWalker->GetParentElement(spPointedElement, &spItemUIElement);
				if (spItemUIElement == nullptr)
					return ;
				CComBSTR	strClassName;
				spItemUIElement->get_CurrentClassName(&strClassName);
				if (strClassName == nullptr || strClassName != L"UIItem")
					return;
			} else 
				spItemUIElement = spPointedElement;

			CComBSTR	strID;
			spItemUIElement->get_CurrentAutomationId(&strID);
			if (strID)
				nIndex = boost::lexical_cast<int>((LPCTSTR)CString(strID));
#endif
		}

		if (nIndex == -1)
			return ;

		CComPtr<IShellView>	spShellView;
		HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
		CComQIPtr<IFolderView>	spFolderView = spShellView;
		if (spFolderView == nullptr)
			return ;

		LPITEMIDLIST pidlChild = nullptr;
		spFolderView->Item(nIndex, &pidlChild);
		if (pidlChild == nullptr)
			return ;

		LPITEMIDLIST pidlFolder = ShellWrap::GetCurIDList(m_spShellBrowser);
		LPITEMIDLIST pidlSelectedItem = ::ILCombine(pidlFolder, pidlChild);

		// リンクは解決する
		CComPtr<IShellItem>	spShellItem;
		hr = ::SHCreateItemFromIDList(pidlSelectedItem, IID_IShellItem, (LPVOID*)&spShellItem);
		if (hr == S_OK) {
			SFGAOF	attribute;
			spShellItem->GetAttributes(SFGAO_LINK, &attribute);
			if (attribute & SFGAO_LINK) {
				LPITEMIDLIST	pidlLink = ShellWrap::GetResolveIDList(pidlSelectedItem);
				if (pidlLink)
					m_wndTabBar.OnTabCreate(pidlLink, false, false, true);
				::ILFree(pidlSelectedItem);
				pidlSelectedItem = nullptr;
			}
		}
		if (pidlSelectedItem) {
			if (IsExistFolderFromIDList(pidlSelectedItem)) {
				m_wndTabBar.OnTabCreate(pidlSelectedItem, false, false, true);
			} else {
				::ILFree(pidlSelectedItem);
			}
		}

		::ILFree(pidlFolder);
		::ILFree(pidlChild);
	}
}










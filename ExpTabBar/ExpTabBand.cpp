// ExpTabBand.cpp : CExpTabBand の実装

#include "stdafx.h"
#include "ExpTabBand.h"
#include <boost/lexical_cast.hpp>
#include "ShellWrap.h"
#include "GdiplusUtil.h"

// コンストラクタ/デストラクタ
CExpTabBand::CExpTabBand() : 
	m_bNavigateCompleted(false), 
	m_wndShellView(this, 1), 
	m_wndListView(this, 2),
	m_wndDirectUI(this, 3),
	m_nIndexTooltip(-1),
	m_bNowTrackMouseLeave(false),
	m_bNowTrackMouseHover(false)
{
	GdiplusInit();
}

CExpTabBand::~CExpTabBand()
{
	GdiplusTerm();

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

		m_wndTabBar.RefreshTab(_T(""));

		HRESULT	hr;
		hr = ::CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&m_spUIAutomation);
		ATLASSERT(m_spUIAutomation);

		/* サムネイルツールチップ作成 */
		m_ThumbnailTooltip.Create(NULL, rcDefault, NULL, WS_POPUP, WS_EX_TOOLWINDOW |  WS_EX_LAYERED | WS_EX_TRANSPARENT);
		ATLASSERT( m_ThumbnailTooltip.IsWindow() );		
		SetLayeredWindowAttributes(m_ThumbnailTooltip.m_hWnd, 0, 255, LWA_ALPHA);
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
		if (m_wndListView.m_hWnd)
			m_wndListView.UnsubclassWindow(TRUE);
		if (m_wndDirectUI.m_hWnd)
			m_wndDirectUI.UnsubclassWindow(TRUE);

		HWND hWnd;
		CComPtr<IShellView>	spShellView;
		HRESULT hr = m_spShellBrowser->QueryActiveShellView(&spShellView);
		if (SUCCEEDED(hr)) {
			hr = spShellView->GetWindow(&hWnd);
			if (SUCCEEDED(hr)) {
				m_wndShellView.SubclassWindow(hWnd);
				m_ListView = ::FindWindowEx(hWnd, NULL, _T("SysListView32"), NULL);
				if (m_ListView.m_hWnd) {
					m_wndListView.SubclassWindow(m_ListView);

					_TrackMouseLeave(m_ListView);
					_TrackMouseHover(m_ListView);
				} else {
					HWND hWndDirectUI = ::FindWindowEx(hWnd, NULL, _T("DirectUIHWND"), NULL);
					if (hWndDirectUI) {
						m_wndDirectUI.SubclassWindow(hWndDirectUI);
						_TrackMouseLeave(hWndDirectUI);
						_TrackMouseHover(hWndDirectUI);
					}
				}
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

LRESULT CExpTabBand::OnListViewGetDispInfo(LPNMHDR pnmh)
{
	SetMsgHandled(FALSE);

	CRect rcItem;
	int nIndex = -1;
	if (m_ListView.m_hWnd) {
		nIndex = _HitTestListView();
		if (nIndex != -1)
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_LABEL);
	} else if (m_wndDirectUI.m_hWnd) {	
		nIndex = _HitTestDirectUI(rcItem);
	}

	if (nIndex != -1 && m_nIndexTooltip != nIndex) {
		m_Tooltip = pnmh->hwndFrom;
		_ShowThumbnailTooltip(nIndex, rcItem);
	}

	return 0;
}

void	CExpTabBand::OnListViewMouseMove(UINT nFlags, CPoint point)
{
	SetMsgHandled(FALSE);

	bool bListView = m_ListView.m_hWnd != NULL;
	if (m_bNowTrackMouseLeave == false)
		_TrackMouseLeave(bListView ? m_ListView.m_hWnd : m_wndDirectUI.m_hWnd);
	if (m_bNowTrackMouseHover == false)
		_TrackMouseHover(bListView ? m_ListView.m_hWnd : m_wndDirectUI.m_hWnd);

	if (m_ThumbnailTooltip.IsWindowVisible()) {
		CPoint ptNow;
		::GetCursorPos(&ptNow);
		if (m_ptLastForMouseMove == ptNow) {
			return ;	// キーボードでの移動
		}
		CRect rcItem;
		int nIndex = bListView ? _HitTestListView() : _HitTestDirectUI(rcItem);
		if (nIndex == -1) {
			_HideThumbnailTooltip();
			return ;
		}

		if (nIndex != -1 && m_nIndexTooltip != nIndex) {
			if (bListView) {
				m_ListView.GetItemRect(nIndex, &rcItem, LVIR_LABEL);
			}
			if (!_ShowThumbnailTooltip(nIndex, rcItem)) {
				_HideThumbnailTooltip();
			}
		}
	}

	//if (m_Tooltip.IsWindow())
	//	m_Tooltip.Activate(TRUE);
}


void	CExpTabBand::OnListViewMouseLeave()
{
	m_bNowTrackMouseLeave = false;

	_HideThumbnailTooltip();
}

void	CExpTabBand::OnListViewMouseHover(WPARAM wParam, CPoint ptPos)
{
	m_bNowTrackMouseHover = false;

	CPoint ptNow;
	::GetCursorPos(&ptNow);
	if (m_ptLastForMouseMove == ptNow) {
		return ;	// キーボードでの移動
	}

	CRect rcItem;
	int nIndex = -1;
	if (m_ListView.m_hWnd) {
		nIndex = _HitTestListView();
		if (nIndex != -1)
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_LABEL);
	} else if (m_wndDirectUI.m_hWnd) {	
		nIndex = _HitTestDirectUI(rcItem);
	}

	if (nIndex != -1 && m_nIndexTooltip != nIndex) {
		_ShowThumbnailTooltip(nIndex, rcItem);
	}
}

void	CExpTabBand::OnListViewKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	CRect rcItem;
	int nIndex = -1;
	bool bListView = m_ListView.m_hWnd != NULL;
	if (bListView) {
		m_wndListView.DefWindowProc();
		nIndex = m_ListView.GetNextItem(-1, LVNI_SELECTED);
		if (nIndex != -1)
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_LABEL);
	} else {
		m_wndDirectUI.DefWindowProc();
		CComPtr<IShellView>	spShellView;
		m_spShellBrowser->QueryActiveShellView(&spShellView);
		CComQIPtr<IFolderView>	spFolderView = spShellView;
		spFolderView->GetFocusedItem(&nIndex);
		if (nIndex != -1)
			rcItem = _GetItemRect(nIndex);
	}

	if (nIndex == -1) {
		_HideThumbnailTooltip();
		return ;
	}
	::GetCursorPos(&m_ptLastForMouseMove);
	m_ptLastForMouseMove;
	if (!_ShowThumbnailTooltip(nIndex, rcItem)) {
		_HideThumbnailTooltip();
	}
}

void	CExpTabBand::OnListViewKillFocus(CWindow wndFocus)
{
	SetMsgHandled(FALSE);
	_HideThumbnailTooltip();
}


/// フォルダーをミドルクリックで新しいタブで開く
void CExpTabBand::OnParentNotify(UINT message, UINT nChildID, LPARAM lParam)
{
	if (message == WM_MBUTTONDOWN) {
		int nIndex = _HitTestListView();
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

int		CExpTabBand::_HitTestDirectUI(CRect& rcItem)
{
	if (m_wndDirectUI.m_hWnd) {
		HRESULT	hr;
		CComPtr<IUIAutomationElement>	spShellViewElement;
		hr = m_spUIAutomation->ElementFromHandle(m_wndShellView, &spShellViewElement);
		if (FAILED(hr))
			return -1;

		CComVariant	vProp(UIA_ListItemControlTypeId);
		CComPtr<IUIAutomationCondition>	spCondition;
		hr = m_spUIAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, vProp, &spCondition);
		if (FAILED(hr)) 
			return -1;

		CComPtr<IUIAutomationElementArray>	spUIElementArray;
		hr = spShellViewElement->FindAll(TreeScope_Descendants, spCondition, &spUIElementArray);
		if (FAILED(hr) || spUIElementArray == nullptr)
			return -1;

		POINT	pt;
		::GetCursorPos(&pt);
		
		int nIndex = -1;
		int nCount = 0;
		spUIElementArray->get_Length(&nCount);
		for (int i = 0; i < nCount; ++i) {
			CComPtr<IUIAutomationElement>	spUIElm;
			spUIElementArray->GetElement(i, &spUIElm);
			BOOL	bOffScreen = FALSE;
			spUIElm->get_CurrentIsOffscreen(&bOffScreen);
			if (bOffScreen)
				continue;
			CRect rcUIItem;
			spUIElm->get_CurrentBoundingRectangle(&rcUIItem);
			if (rcUIItem.PtInRect(pt)) {	// 見つかった
				nIndex = i;

				CComBSTR strItemNameDisplay = L"System.ItemNameDisplay";
				CComVariant	vProp(strItemNameDisplay);
				CComPtr<IUIAutomationCondition>	spCondition1;
				m_spUIAutomation->CreatePropertyCondition(UIA_AutomationIdPropertyId, vProp, &spCondition1);
				if (spCondition1) {
					CComPtr<IUIAutomationElement>	spUIItenNameElm;
					spUIElm->FindFirst(TreeScope_Children, spCondition1, &spUIItenNameElm);
					if (spUIItenNameElm) {
						spUIItenNameElm->get_CurrentBoundingRectangle(&rcItem);
						m_wndDirectUI.ScreenToClient(&rcItem);
					}
				}
				break;
			}
		}
		return nIndex;
	}
	return -1;
}

int		CExpTabBand::_HitTestListView(const CPoint& pt)
{
	if (m_ListView.m_hWnd) {
		UINT Flags;
		int nIndex = m_ListView.HitTest(pt, &Flags);
		return nIndex;
	} else if (m_wndDirectUI.m_hWnd) {
		CRect rcItem;
		return _HitTestDirectUI(rcItem);
	}
	return -1;
}

int		CExpTabBand::_HitTestListView()
{
	CPoint pt;
	::GetCursorPos(&pt);
	if (m_ListView.m_hWnd) {		
		m_ListView.ScreenToClient(&pt);
	} else if (m_wndDirectUI.m_hWnd) {
		m_wndDirectUI.ScreenToClient(&pt);
	} else {
		return -1;
	}

	return _HitTestListView(pt);

}

CRect	CExpTabBand::_GetItemRect(int nIndex)
{
	CRect rcItem;
	HRESULT	hr;
	CComPtr<IUIAutomationElement>	spShellViewElement;
	hr = m_spUIAutomation->ElementFromHandle(m_wndShellView, &spShellViewElement);
	if (FAILED(hr))
		return rcItem;

	CComBSTR strIndex(boost::lexical_cast<std::wstring>(nIndex).c_str());
	CComVariant	vProp(strIndex);
	CComPtr<IUIAutomationCondition>	spCondition;
	hr = m_spUIAutomation->CreatePropertyCondition(UIA_AutomationIdPropertyId, vProp, &spCondition);
	if (FAILED(hr)) 
		return rcItem;

	CComPtr<IUIAutomationElement>	spUIElement;
	hr = spShellViewElement->FindFirst(TreeScope_Descendants, spCondition, &spUIElement);
	if (FAILED(hr) || spUIElement == nullptr)
		return rcItem;

	CComBSTR strItemNameDisplay = L"System.ItemNameDisplay";
	CComVariant	vProp2(strItemNameDisplay);
	CComPtr<IUIAutomationCondition>	spCondition2;
	m_spUIAutomation->CreatePropertyCondition(UIA_AutomationIdPropertyId, vProp2, &spCondition2);
	if (spCondition2) {
		CComPtr<IUIAutomationElement>	spUIItenNameElm;
		spUIElement->FindFirst(TreeScope_Children, spCondition2, &spUIItenNameElm);
		if (spUIItenNameElm) {
			spUIItenNameElm->get_CurrentBoundingRectangle(&rcItem);
			m_wndDirectUI.ScreenToClient(&rcItem);
			return rcItem;
		}
	}
	return rcItem;
}

bool	CExpTabBand::_ShowThumbnailTooltip(int nIndex, CRect rcItem)
{
	//SetLayeredWindowAttributes(m_pThumbnailTooltip->m_hWnd, 0, 255, LWA_ALPHA);
	//m_pThumbnailTooltip->ShowWindow(TRUE);

	LPITEMIDLIST pidl = ShellWrap::GetIDListByIndex(m_spShellBrowser, nIndex);
	if (pidl == nullptr)
		return false;

	CString path = ShellWrap::GetFullPathFromIDList(pidl);
	::ILFree(pidl);
	CString strExt = Misc::GetPathExtention(path);
	strExt.MakeLower();
	if (strExt == _T("jpg") || strExt == _T("jpeg") || strExt == _T("png") || strExt == _T("gif") || strExt == _T("bmp"))
	{
		m_nIndexTooltip = nIndex;
		m_wndShellView.ClientToScreen(&rcItem);
		bool bShow = m_ThumbnailTooltip.ShowThumbnailTooltip((LPCWSTR)path, rcItem);
		if (bShow) {
			if (m_Tooltip.IsWindow()) {
				m_Tooltip.Activate(FALSE);
			}
		}
		return bShow;
	}
	return false;
}

void	CExpTabBand::_HideThumbnailTooltip()
{
	m_ThumbnailTooltip.HideThumbnailTooltip();
	m_nIndexTooltip = -1;
	if (m_Tooltip.IsWindow())
		m_Tooltip.Activate(TRUE);
}

void	CExpTabBand::_TrackMouseLeave(HWND hWnd)
{
	TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT) };
	tme.hwndTrack = hWnd;
	tme.dwFlags	= TME_LEAVE;
	::TrackMouseEvent(&tme);
	m_bNowTrackMouseLeave = true;
}

void	CExpTabBand::_TrackMouseHover(HWND hWnd)
{
	TRACKMOUSEEVENT tme = { sizeof(TRACKMOUSEEVENT) };
	tme.hwndTrack = hWnd;
	tme.dwFlags	= TME_HOVER;
	tme.dwHoverTime	= HOVER_DEFAULT;
	::TrackMouseEvent(&tme);
	m_bNowTrackMouseHover = true;
}



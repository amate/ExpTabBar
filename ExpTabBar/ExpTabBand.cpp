// ExpTabBand.cpp : CExpTabBand の実装

#include "stdafx.h"
#include "ExpTabBand.h"
#include <boost/lexical_cast.hpp>
#include <TlHelp32.h>
#include "ShellWrap.h"
#include "GdiplusUtil.h"
#include "Logger.h"

namespace {


	bool	IsImageFile(const CString& path)
	{
		if (::PathIsDirectory(path))
			return false;

		CString strExt = Misc::GetPathExtention(path);
		strExt.MakeLower();
		if (strExt == _T("webp") || strExt == _T("jpg") || strExt == _T("jpeg") || strExt == _T("png") || strExt == _T("gif") || strExt == _T("bmp")) {
			return true;
		} else {
			return false;
		}
	}

}	// namespace

// コンストラクタ/デストラクタ
CExpTabBand::CExpTabBand() : 
	m_bNavigateCompleted(false), 
	m_wndShellView(this, 1), 
	m_wndListView(this, 2),
	m_wndDirectUI(this, 3),
	m_wndExplorer(this, 6),
	m_wndAddressBarProgress(this, 7),
	m_wndAddressBarEditCtrl(this, 8),
	m_wndShellTabWindow(this, 9),
	m_nIndexTooltip(-1),
	m_bNowTrackMouseLeave(false),
	m_bNowTrackMouseHover(false),
	m_bWheelThumbnailView(false)
{
}

CExpTabBand::~CExpTabBand()
{
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
	if (fShow == FALSE)	// 表示されるときはエクスプローラーにフォーカスが当たってるので登録しなくていい
		_RegisterExecuteCommandVerb(false);

	if (fShow && m_wndExplorer.IsWindow() == FALSE && m_wndTabBar.m_hWnd)
		m_wndExplorer.SubclassWindow(m_wndTabBar.GetTopLevelWindow());
	else if (fShow == FALSE && m_wndExplorer.IsWindow())
		m_wndExplorer.UnsubclassWindow();

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
	INFO_LOG << L"CExpTabBand::SetSite : punkSite : " << punkSite;
	HRESULT hr;
	if (punkSite) {
		//Get the parent window.
		CComQIPtr<IOleWindow> pOleWindow(punkSite);
		if (pOleWindow == nullptr) {
			INFO_LOG << L"CExpTabBand::SetSite : pOleWindow faild";
			return E_FAIL;
		}

		HWND	hWndParent;
		pOleWindow->GetWindow(&hWndParent);
		if (hWndParent == NULL) {
			ERROR_LOG << L"CExpTabBand::SetSite : pOleWindow->GetWindow failed";
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
				ERROR_LOG << L"CExpTabBand::SetSite : DispEventAdvise(m_spWebBrowser2) failed";
				ATLASSERT(FALSE);
				return E_FAIL;
			}
		}

		pSP->QueryService(SID_SShellBrowser, &m_spShellBrowser);
		ATLASSERT(m_spShellBrowser);
		if (m_spShellBrowser == nullptr) {
			ATLASSERT(FALSE);
			ERROR_LOG << L"CExpTabBand::SetSite : m_spShellBrowser failed";
			return E_FAIL;
		}

		GdiplusInit();

		m_wndTabBar.Initialize(punkSite, this);

		/* タブバー作成 */
		m_wndTabBar.Create(hWndParent);
		if (m_wndTabBar.IsWindow() == FALSE) {
			ATLASSERT(FALSE);
			ERROR_LOG << L"CExpTabBand::SetSite : m_wndTabBar.Create(hWndParent) failed";
			return E_FAIL;
		}

		m_wndTabBar.RefreshTab(_T(""));

		HRESULT	hr;
		hr = ::CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&m_spUIAutomation);
		ATLASSERT(m_spUIAutomation);

		/* サムネイルツールチップ作成 */
		m_ThumbnailTooltip.Create(NULL, rcDefault, NULL, WS_POPUP, 
			WS_EX_TOOLWINDOW |  WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOPMOST);
		ATLASSERT( m_ThumbnailTooltip.IsWindow() );		
		SetLayeredWindowAttributes(m_ThumbnailTooltip.m_hWnd, 0, 255, LWA_ALPHA);
		m_ThumbnailTooltip.SetIsFolderDetailViewFunc([this]() -> bool {
			CComPtr<IShellView>	spShellView;
			m_spShellBrowser->QueryActiveShellView(&spShellView);
			CComQIPtr<IFolderView>	spFolderView = spShellView;
			UINT viewMode = 0;
			spFolderView->GetCurrentViewMode(&viewMode);
			if (viewMode == FVM_ICON) {
				return false;
			} else {
				return true;
			}
		});

		m_wndShellTabWindow.SubclassWindow(CWindow(hWndParent).GetParent().GetParent());

		m_wndExplorer.SubclassWindow(CWindow(hWndParent).GetTopLevelWindow());
		if (m_wndExplorer == ::GetFocus() || m_wndExplorer.IsChild(::GetFocus()))
			_RegisterExecuteCommandVerb(false);

		_SubclassAddressBarProgress();

	} else {
		DispEventUnadvise(m_spWebBrowser2);

		if (m_wndAddressBarProgress)
			m_wndAddressBarProgress.UnsubclassWindow(TRUE);
		if (m_wndAddressBarEditCtrl)
			m_wndAddressBarEditCtrl.UnsubclassWindow(TRUE);

		if (m_wndShellView)
			m_wndShellView.UnsubclassWindow(TRUE);
		if (m_wndListView)
			m_wndListView.UnsubclassWindow(TRUE);
		if (m_wndDirectUI)
			m_wndDirectUI.UnsubclassWindow(TRUE);

		if (m_wndShellTabWindow)
			m_wndShellTabWindow.UnsubclassWindow(TRUE);
		if (m_wndExplorer)
			m_wndExplorer.UnsubclassWindow(TRUE);

		m_ThumbnailTooltip.DestroyWindow();

		m_spUIAutomation.Release();

		m_wndTabBar.DestroyWindow();
		m_wndTabBar.UnInitialize();

		GdiplusTerm();

		m_spShellBrowser.Release();
		m_spWebBrowser2.Release();

		_RegisterExecuteCommandVerb(false);
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
	if (m_spShellBrowser == nullptr) {
		ATLASSERT(FALSE);
		return;
	}

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

					if (CTabSkinDarkTheme::IsDarkMode()) {
						m_ListView.SetTextColor(0xFFFFFF);
					}
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
	_SetNoFullRowSelect();
}


void CExpTabBand::OnTitleChange(BSTR title)
{
	m_wndTabBar.RefreshTab(title);
	m_ThumbnailTooltip.OnLocationChanged();
}

// Message map

/// 選択されているアイテムが変わった
LRESULT CExpTabBand::OnListViewItemChanged(LPNMHDR pnmh)
{
	m_wndShellView.DefWindowProc();

	LPNMLISTVIEW	pnmlv = (LPNMLISTVIEW)pnmh;
	if (pnmlv->iItem == -1)
		return 0;

	// ドラッグ選択するとサムネイルツールチップは消すようにする
	if (::GetCapture())
		_HideThumbnailTooltip();
	
	/* アイコン部分が再描写されないバグに対処 */
	RECT rcItem;
	m_ListView.GetItemRect(pnmlv->iItem, &rcItem, LVIR_ICON);
	m_ListView.InvalidateRect(&rcItem);
	int nIndex = m_ListView.GetNextItem(-1, LVIS_SELECTED);
	do {
		m_ListView.GetItemRect(nIndex, &rcItem, LVIR_ICON);
		m_ListView.InvalidateRect(&rcItem);
		nIndex = m_ListView.GetNextItem(nIndex, LVIS_SELECTED);
	} while (nIndex != -1);

	return 0;
}

LRESULT CExpTabBand::OnListViewGetDispInfo(LPNMHDR pnmh)
{
	if (m_bWheelThumbnailView)
		return 0;

	SetMsgHandled(FALSE);

	CRect rcItem;
	int nIndex = -1;
	if (m_ListView.m_hWnd) {
		nIndex = _HitTestListView();
		if (nIndex != -1)
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_SELECTBOUNDS);
	} else if (m_wndDirectUI.m_hWnd) {	
		nIndex = _HitTestDirectUI(rcItem);
	}


	m_Tooltip = pnmh->hwndFrom;
	if (nIndex != -1 && m_nIndexTooltip != nIndex) {
		_ShowThumbnailTooltip(nIndex, rcItem);
	}

	return 0;
}

void	CExpTabBand::OnListViewMouseMove(UINT nFlags, CPoint point)
{
	SetMsgHandled(FALSE);
	if (m_bWheelThumbnailView)
		return;

	bool bListView = m_ListView.m_hWnd != NULL;
	if (m_bNowTrackMouseLeave == false)
		_TrackMouseLeave(bListView ? m_ListView.m_hWnd : m_wndDirectUI.m_hWnd);
	if (m_bNowTrackMouseHover == false)
		_TrackMouseHover(bListView ? m_ListView.m_hWnd : m_wndDirectUI.m_hWnd);

	if (m_ThumbnailTooltip.IsWindowVisible()) {
		CPoint ptNow;
		::GetCursorPos(&ptNow);
		enum { kAreaSize = 10 };
		CRect rcMouseArea(CPoint(m_ptLastForMouseMove.x - (kAreaSize / 2), m_ptLastForMouseMove.y - (kAreaSize / 2)), CSize(kAreaSize, kAreaSize));
		if (rcMouseArea.PtInRect(ptNow)) {
			return ;	// キーボードでの移動
		}
		CRect rcItem;
		int nIndex = bListView ? _HitTestListView() : _HitTestDirectUI(rcItem);
		if (nIndex == -1) {
			m_wndTabBar.SetTimer(kHideThumbnailTooltipTimerID, kHideThumbnailTooltipTimerInterval);
			//_HideThumbnailTooltip();
			return ;
		}

		if (nIndex != -1 && m_nIndexTooltip != nIndex) {
			if (bListView) {
				m_ListView.GetItemRect(nIndex, &rcItem, LVIR_SELECTBOUNDS);
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

	if (m_Tooltip.IsWindow() == FALSE)	// 初回は何もしない
		return ;

	CRect rcItem;
	int nIndex = -1;
	if (m_ListView.m_hWnd) {
		nIndex = _HitTestListView();
		if (nIndex != -1)
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_SELECTBOUNDS);
	} else if (m_wndDirectUI.m_hWnd) {	
		nIndex = _HitTestDirectUI(rcItem);
	}

	if (nIndex != -1 && m_nIndexTooltip != nIndex) {
		_ShowThumbnailTooltip(nIndex, rcItem);
	}
}

void	CExpTabBand::OnListViewKeyUp(UINT nChar, UINT nRepCnt, UINT nFlags)
{
//	ATLTRACE(_T("RepCnt : %d\n"), nRepCnt);

	if (nRepCnt > 1) {
		SetMsgHandled(FALSE);
		return;
	}

	switch (nChar) {
	case VK_UP:
	case VK_DOWN:
	case VK_LEFT:
	case VK_RIGHT:
	case VK_HOME:
	case VK_END:
	case VK_PRIOR:
	case VK_NEXT:
		break;

	case VK_ESCAPE:
		SetMsgHandled(FALSE);
		_HideThumbnailTooltip();
		return;

	default:
		SetMsgHandled(FALSE);
		return;
	}

	CRect rcItem;
	int nIndex = -1;
	bool bListView = m_ListView.m_hWnd != NULL;
	if (bListView) {
		int nPrevIndex = m_ListView.GetNextItem(-1, LVNI_SELECTED);
		m_wndListView.DefWindowProc();
		nIndex = m_ListView.GetNextItem(-1, LVNI_SELECTED);
		if (nIndex != -1) {
			if (nPrevIndex == nIndex) {
				if (nChar == VK_RIGHT)
					++nIndex;
				else if (nChar == VK_LEFT)
					--nIndex;
				if (0 <= nIndex && nIndex < m_ListView.GetItemCount()) {
					m_ListView.SelectItem(nIndex);
				}
			}
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_SELECTBOUNDS);
		}
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
	if (m_ThumbnailTooltip.IsShowThumbnailTooltip()) {
		::GetCursorPos(&m_ptLastForMouseMove);
		if (!_ShowThumbnailTooltip(nIndex, rcItem)) {
			_HideThumbnailTooltip();
		}
	}
}

void	CExpTabBand::OnListViewKillFocus(CWindow wndFocus)
{
	SetMsgHandled(FALSE);
	_HideThumbnailTooltip();
}

BOOL	CExpTabBand::OnListViewMouseWheel(UINT nFlags, short zDelta, CPoint pt)
{
	if (nFlags == MK_XBUTTON1) {
		int nIndex = -1;
		int nItemCount = 0;
		bool bListView = m_ListView.m_hWnd != NULL;
		if (bListView) {
			nIndex = m_ListView.GetNextItem(-1, LVNI_SELECTED);
			nItemCount = m_ListView.GetItemCount();
		} else {
			CComPtr<IShellView>	spShellView;
			m_spShellBrowser->QueryActiveShellView(&spShellView);
			CComQIPtr<IFolderView>	spFolderView = spShellView;
			spFolderView->GetFocusedItem(&nIndex);
			spFolderView->ItemCount(SVGIO_ALLVIEW, &nItemCount);
		}
		if (nIndex == -1) {
			SetMsgHandled(FALSE);
			return 0;
		}

		CRect rcItem;
		int nNextIndex = zDelta > 0 ? nIndex - 1 : nIndex + 1;
		if (nNextIndex == -1 || nNextIndex == nItemCount) {
			SetMsgHandled(FALSE);
			return 0;
		}
		
		if (bListView) {
			m_ListView.SelectItem(nNextIndex);
			m_ListView.GetItemRect(nIndex, &rcItem, LVIR_SELECTBOUNDS);
		} else {
			CComPtr<IShellView>	spShellView;
			m_spShellBrowser->QueryActiveShellView(&spShellView);
			CComQIPtr<IFolderView>	spFolderView = spShellView;
			spFolderView->SelectItem(nNextIndex, SVSI_FOCUSED | SVSI_ENSUREVISIBLE);
			rcItem = _GetItemRect(nIndex);
		}
		m_bWheelThumbnailView = true;
		::GetCursorPos(&m_ptLastForMouseMove);
		if (!_ShowThumbnailTooltip(nNextIndex, rcItem, true)) {
			_HideThumbnailTooltip();
		}
	} else {
		SetMsgHandled(FALSE);
	}
	return 0;
}

void	CExpTabBand::OnListViewXButtonUp(int fwButton, int dwKeys, CPoint ptPos)
{
	if (fwButton == XBUTTON1) {
		if (m_bWheelThumbnailView) {
			m_bWheelThumbnailView = false;
			return;
		}
	}
	SetMsgHandled(FALSE);
}


/// フォルダーをミドルクリックで新しいタブで開く
void CExpTabBand::OnParentNotify(UINT message, UINT nChildID, LPARAM lParam)
{
	if (message == WM_MBUTTONDOWN) {
		int nIndex = _HitTestListView();
		if (nIndex == -1) {
			if (m_ThumbnailTooltip.IsShowThumbnailTooltip()) {
				_HideThumbnailTooltip();
			}
			return;
		}

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
		CString strFullPath = ShellWrap::GetFullPathFromIDList(pidlSelectedItem);
		if (Misc::GetPathExtention(strFullPath).CompareNoCase(_T("lnk")) == 0) {
			LPITEMIDLIST	pidlLink = ShellWrap::GetResolveIDList(pidlSelectedItem);
			if (pidlLink) {
				m_wndTabBar.OnTabCreate(pidlLink, false, false, true);
				::CoTaskMemFree(pidlSelectedItem);
				pidlSelectedItem = nullptr;
			}
		} else {
			if (m_ThumbnailTooltip.IsShowThumbnailTooltip() == false && IsImageFile(strFullPath)) {
				CRect rcItem;
				if (m_ListView.m_hWnd) {
					m_ListView.GetItemRect(nIndex, &rcItem, LVIR_SELECTBOUNDS);
				} else if (m_wndDirectUI.m_hWnd) {
					ATLVERIFY(_HitTestDirectUI(rcItem) == nIndex);
				}
				_ShowThumbnailTooltip(nIndex, rcItem, true);
			} else {
				_HideThumbnailTooltip();
			}
		}


		if (pidlSelectedItem) {
			if (IsExistFolderFromIDList(pidlSelectedItem)) {
				m_wndTabBar.OnTabCreate(pidlSelectedItem, false, false, true);
			} else {
				::CoTaskMemFree(pidlSelectedItem);
			}
		}

		::CoTaskMemFree(pidlFolder);
		::CoTaskMemFree(pidlChild);
	}
}


// タブバーダブルクリック
// Ctrlを押しながらでフォルダ内のすべての画像のサムネイルキャッシュを作成
// Shiftを押しながらでサムネイルキャッシュを削除
void	CExpTabBand::OnTabBarLButtonDblClk(UINT nFlags, CPoint point)
{
	if (::GetKeyState(VK_CONTROL) < 0) {
		_SetNoFullRowSelect();
#if 0
		std::vector<CString> vec;
		CComPtr<IShellView>	spShellView;
		m_spShellBrowser->QueryActiveShellView(&spShellView);
		CComQIPtr<IFolderView>	spFolderView = spShellView;
		if (spFolderView) {
			int nCount = 0;
			spFolderView->ItemCount(SVGIO_ALLVIEW, &nCount);
			LPITEMIDLIST pidlFolder = ShellWrap::GetCurIDList(m_spShellBrowser);
			for (int i = 0; i < nCount; ++i) {
				LPITEMIDLIST pidlChild = nullptr;
				spFolderView->Item(i, &pidlChild);
				LPITEMIDLIST pidl = ::ILCombine(pidlFolder, pidlChild);
				CString strPath = ShellWrap::GetFullPathFromIDList(pidl);
				if (IsImageFile(strPath)) {
					vec.push_back(strPath);
				}

				::CoTaskMemFree(pidl);
				::CoTaskMemFree(pidlChild);
			}
			::CoTaskMemFree(pidlFolder);

			m_ThumbnailTooltip.LockImageCache();
			//std::thread	td([this, vec]() {
				for (auto it = vec.begin(); it != vec.end(); ++it) {
					if (::GetAsyncKeyState(VK_ESCAPE) < 0)
						break;
					m_ThumbnailTooltip.AddThumbnailCache((LPCTSTR)*it);
				}
			//});
			//td.detach();
		}
#endif
	} else if(::GetKeyState(VK_SHIFT) < 0) {
		m_ThumbnailTooltip.ClearImageCache();
		m_wndTabBar.MessageBox(L"サムネイルキャッシュをクリアしました。", L"確認", MB_OK);
	}
}


void	CExpTabBand::OnTabBarTimer(UINT_PTR nIDEvent)
{
	if (nIDEvent == kHideThumbnailTooltipTimerID) {
		bool bListView = m_ListView.m_hWnd != NULL;
		if (m_ThumbnailTooltip.IsWindowVisible()) {
			CRect rcItem;
			int nIndex = bListView ? _HitTestListView() : _HitTestDirectUI(rcItem);
			if (nIndex == -1) {
				_HideThumbnailTooltip();
			}
		}
		m_wndTabBar.KillTimer(kHideThumbnailTooltipTimerID);
	} else {
		SetMsgHandled(FALSE);
	}
}


void	CExpTabBand::OnExplorerActivate(UINT nState, BOOL bMinimized, CWindow wndOther)
{
	SetMsgHandled(FALSE);

	if (nState == WA_INACTIVE)
		_HideThumbnailTooltip();

	_RegisterExecuteCommandVerb(!(nState != WA_INACTIVE));
}

/// なかなか終了しないエクスプローラーを強制終了させる
void	CExpTabBand::OnExplorerDestroy()
{
	SetMsgHandled(FALSE);
#if 0
	OSVERSIONINFO	osvi = { sizeof(osvi) };
	GetVersionEx(&osvi);
	if (osvi.dwMajorVersion == 6 && osvi.dwMinorVersion == 1) {	// Win7
		std::thread([]{
			::Sleep(5 * 1000);

			DWORD processID = GetCurrentProcessId();
			HANDLE h = ::OpenProcess(PROCESS_TERMINATE, FALSE, processID);
			ATLASSERT(h);
			::TerminateProcess(h, 0);
			::CloseHandle(h);
		}).detach();
	}
#endif
}

void	CExpTabBand::OnAddressBarProgressParentNotify(UINT message, UINT nChildID, LPARAM lParam)
{
	m_wndAddressBarProgress.DefWindowProc();
	if (message == WM_CREATE && nChildID == 0xA205) {	// ConboBoxEx32
		_SubclassAddressBarEditCtrl();
	}
}

void	CExpTabBand::OnAddressBarEditKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	if (nChar == VK_RETURN && CTabBarConfig::s_bAddressBarNewTabOpen) {
		CString strPath;
		m_wndAddressBarEditCtrl.GetWindowText(strPath.GetBuffer(MAX_PATH), MAX_PATH);
		strPath.ReleaseBuffer();
		ATLTRACE(_T("OnAddressBarEditKeyDown : %s\n"), (LPCTSTR)strPath);
		LPITEMIDLIST pidl = ShellWrap::CreateIDListFromFullPath(strPath);
		if (pidl && ShellWrap::IsExistFolderFromIDList(pidl)) {
			m_wndTabBar.ExternalOpen(pidl);
			return ;
		}
	}
	m_wndAddressBarEditCtrl.DefWindowProc();
}

void	CExpTabBand::OnChildActivate()
{
	m_wndTabBar.TopTabActivate();
}

// private:

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
		for (int i = 0; i < nCount ; ++i) {
			CComPtr<IUIAutomationElement>	spUIElm;
			spUIElementArray->GetElement(i, &spUIElm);
			BOOL	bOffScreen = FALSE;
			spUIElm->get_CurrentIsOffscreen(&bOffScreen);
			if (bOffScreen)
				continue;
			CRect rcUIItem;
			spUIElm->get_CurrentBoundingRectangle(&rcUIItem);
			if (rcUIItem.PtInRect(pt)) {	// 見つかった
				CComBSTR strId;
				if (SUCCEEDED(spUIElm->get_CurrentAutomationId(&strId)) && strId)
					nIndex = ::_wtoi(strId);
				else
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

bool	CExpTabBand::_ShowThumbnailTooltip(int nIndex, CRect rcItem, bool bForceShow /*= false*/)
{
	if (!CThumbnailTooltipConfig::s_bUseThumbnailTooltip) {
		return false;
	}
	//SetLayeredWindowAttributes(m_pThumbnailTooltip->m_hWnd, 0, 255, LWA_ALPHA);
	//m_pThumbnailTooltip->ShowWindow(TRUE);

	LPITEMIDLIST pidl = ShellWrap::GetIDListByIndex(m_spShellBrowser, nIndex);
	if (pidl == nullptr)
		return false;

	// Altキーを押しているときだけサムネイルを表示する
	if (bForceShow == false && CThumbnailTooltipConfig::s_bShowThumbnailOnAlt && (GetKeyState(VK_MENU) < 0) == false) {
		::CoTaskMemFree(pidl);
		return false;
	}

	CString path = ShellWrap::GetFullPathFromIDList(pidl);
	::CoTaskMemFree(pidl);
	if (IsImageFile(path)) {
		m_nIndexTooltip = nIndex;
		m_wndShellView.ClientToScreen(&rcItem);
		bool bShow = m_ThumbnailTooltip.ShowThumbnailTooltip((LPCWSTR)path, rcItem);
		if (bShow) {
			if (m_Tooltip.IsWindow()) {
				m_Tooltip.Activate(FALSE);
			}

			{	// 周辺の画像をキャッシュに加える
				auto funcAddCacheSurroundImage = [=](bool bPlus) {
					//enum { kMaxDistance = 10 };
					for (int i = 1; i <= CThumbnailTooltipConfig::s_nMaxPreCache; ++i) {
						int nNextIndex = nIndex;
						nNextIndex += bPlus ? +i : -i;
						LPITEMIDLIST pidl = ShellWrap::GetIDListByIndex(m_spShellBrowser, nNextIndex);
						if (pidl == nullptr)
							return;

						CString path = ShellWrap::GetFullPathFromIDList(pidl);
						::CoTaskMemFree(pidl);
						if (IsImageFile(path))
							m_ThumbnailTooltip.AddThumbnailCache((LPCWSTR)path);
					}
				};
				funcAddCacheSurroundImage(true);
				funcAddCacheSurroundImage(false);
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
	if (m_Tooltip.IsWindow()) {
		m_Tooltip.Activate(TRUE);
	}

	if (m_bWheelThumbnailView) {
		m_bWheelThumbnailView = false;
	}
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

/// 設定されていれば一列選択表示をやめる
void	CExpTabBand::_SetNoFullRowSelect()
{
	//if (CTabBarConfig::s_bNoFullRowSelect == false) {
	//	return ;
	//}
	bool bNoFullRaw = CTabBarConfig::s_bNoFullRowSelect;
	HRESULT	hr = S_OK;
	//LPITEMIDLIST	pidl = ShellWrap::GetCurIDList(m_spShellBrowser);
	//if (pidl) {
	//	CString strURL = ShellWrap::GetFullPathFromIDList(pidl);
	//	// "コンピューター"は一列表示にする
	//	if (strURL == _T("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"))
	//		bNoFullRaw = false;
	//	::CoTaskMemFree(pidl);
	//}
	// "コンピューター"は一列表示にする
	//if (strURL == _T("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"))
	//	bNoFullRaw = false;

	CComQIPtr<IFolderViewOptions> spFolderOptions = m_spShellBrowser;
	if (spFolderOptions) {
		if (bNoFullRaw) {
			hr = spFolderOptions->SetFolderViewOptions(FVO_VISTALAYOUT, FVO_VISTALAYOUT);
		} else {
			hr = spFolderOptions->SetFolderViewOptions(FVO_VISTALAYOUT, FVO_DEFAULT);
		}
	}

	CComPtr<IDispatch>	spDisp;
	m_spWebBrowser2->get_Document(&spDisp);
	CComQIPtr<IShellFolderViewDual3>	spShellFolderView3 = spDisp;
	if (spShellFolderView3) {
		DWORD	dwFlags = 0;
		spShellFolderView3->get_FolderFlags(&dwFlags);
		if (CTabBarConfig::s_bAlwaysShowColumHeaders)
			dwFlags &= ~FWF_NOHEADERINALLVIEWS;	// Only show the column header in details view mode.　を取り除く
		if (bNoFullRaw)
			dwFlags &= ~FWF_FULLROWSELECT;
		else
			dwFlags |= FWF_FULLROWSELECT;
		spShellFolderView3->put_FolderFlags(dwFlags);
	}
}


void	CExpTabBand::_RegisterExecuteCommandVerb(bool bRegister)
{
#if 0
	OSVERSIONINFO	osvi = { sizeof(osvi) };
	GetVersionEx(&osvi);
	if ( !(osvi.dwMajorVersion == 6 && osvi.dwMinorVersion == 1) )	// Win7以外
		return ;

	CRegKey	rkFolderCommand;
	if (bRegister) {
		if (::GetAsyncKeyState(VK_MENU) < 0)	// alt 押しながらエクスプローラーをinactiveにするときは登録しない
			return;
		if (rkFolderCommand.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Folder\\shell\\open\\command")) == ERROR_SUCCESS) {
			rkFolderCommand.SetStringValue(_T("DelegateExecute"), OPENINPROCESSGUID);
		}
	} else {
		if (rkFolderCommand.Open(HKEY_CURRENT_USER, _T("Software\\Classes\\Folder\\shell")) == ERROR_SUCCESS)
			rkFolderCommand.RecurseDeleteKey(_T("open"));
	}
#endif
}

bool	CExpTabBand::_SubclassAddressBarProgress()
{
	LPCTSTR targetClassLoad[] = {
		_T("WorkerW"), _T("ReBarWindow32"), _T("Address Band Root"), _T("msctls_progress32"), nullptr
	};

	std::function<HWND (HWND, int)>	funcFindWindow;
	funcFindWindow = [&](HWND hWnd, int nIndex) -> HWND {
		if (targetClassLoad[nIndex] == nullptr)
			return hWnd;
		HWND hWndFound = FindWindowEx(hWnd, NULL, targetClassLoad[nIndex], NULL);
		if (hWndFound) {
			return funcFindWindow(hWndFound, nIndex + 1);
		}
		return NULL;
	};
	HWND hWndProgress = funcFindWindow(m_wndExplorer, 0);
	//ATLASSERT( hWndProgress );
	if (hWndProgress) {
		m_wndAddressBarProgress.SubclassWindow(hWndProgress);
		return true;
	}
	return false;
}

bool	CExpTabBand::_SubclassAddressBarEditCtrl()
{
	LPCTSTR targetClassLoad[] = {
		_T("ComboBoxEx32"), _T("ComboBox"), _T("Edit"), nullptr
	};
	std::function<HWND (HWND, int)>	funcFindWindow;
	funcFindWindow = [&](HWND hWnd, int nIndex) -> HWND {
		if (targetClassLoad[nIndex] == nullptr)
			return hWnd;
		HWND hWndFound = FindWindowEx(hWnd, NULL, targetClassLoad[nIndex], NULL);
		if (hWndFound) {
			return funcFindWindow(hWndFound, nIndex + 1);
		}
		return NULL;
	};
	HWND hWndEdit = funcFindWindow(m_wndAddressBarProgress, 0);
	ATLASSERT( hWndEdit );
	if (hWndEdit) {
		m_wndAddressBarEditCtrl.SubclassWindow(hWndEdit);
		return true;
	}
	return false;
}






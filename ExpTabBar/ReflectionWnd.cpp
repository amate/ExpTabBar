// ReflectionWnd.cpp

#include "stdafx.h"
#include "ReflectionWnd.h"


/////////////////////////////////////////////////////////////////////////////
// CReflectionWnd

// コンストラクタ/デストラクタ
CReflectionWnd::CReflectionWnd()
{
}

CReflectionWnd::~CReflectionWnd()
{
#if 1
	if (m_wndTabBar.IsWindow()) {
		m_wndTabBar.DestroyWindow();
	}
#endif
	if (IsWindow())	{
		DestroyWindow();
	}
	::UnhookWindowsHookEx( g_hHook );
}


LRESULT CReflectionWnd::OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
{
	RECT rect;
	GetClientRect(&rect);
	m_wndTabBar.Create(m_hWnd, rect, _T("DonutTabBar"), WS_CHILD | WS_VISIBLE);

	g_hWndTabBar = m_wndTabBar.m_hWnd;	// 通知先を設定
	g_hHook = ::SetWindowsHookEx(WH_MOUSE, MouseProc, NULL/*g_hInst*/, GetCurrentThreadId());

	return 0;
}

HHOOK	g_hHook;
HWND	g_hWndTabBar;

// --------------------------------------------------------
//    フックプロシージャ
// --------------------------------------------------------

LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode < 0) {
		return CallNextHookEx(g_hHook, nCode, wParam, lParam);
	}
	LPMOUSEHOOKSTRUCT	pmh = (LPMOUSEHOOKSTRUCT)lParam;
	switch (wParam) {
	case WM_MBUTTONDOWN:
		ATLTRACE(_T("WM_MBUTTONDOWN\n"));
		::SendMessage(g_hWndTabBar, WM_NEWTABBUTTON, 0, (LPARAM)&pmh->pt);
		return TRUE;
		break;
	}

	return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}
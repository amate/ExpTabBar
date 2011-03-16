// ReflectionWnd.h

#pragma once

#include "DonutTabBar.h"

extern HHOOK	g_hHook;
extern HWND		g_hWndTabBar;

LRESULT CALLBACK MouseProc(int nCode, WPARAM wParam, LPARAM lParam);


/////////////////////////////////////////////////////////////////////////////
// CReflectionWnd
class CReflectionWnd : public CWindowImpl<CReflectionWnd>
{
public:
	DECLARE_WND_CLASS(NULL)

	// data members
	HWND			m_hWndTop;			

	// コンストラクタ/デストラクタ
	CReflectionWnd();
	virtual ~CReflectionWnd();


	BEGIN_MSG_MAP(CReflectionWnd)
		MESSAGE_HANDLER(WM_CREATE, OnCreate)
		CHAIN_MSG_MAP_MEMBER(m_wndTabBar)
	END_MSG_MAP()

// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

	LRESULT OnCreate(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);

	// タブバーの関数を呼び出したいときはこれを使う
	CDonutTabBar& GetTabBar() { return m_wndTabBar; };

private:
	CDonutTabBar	m_wndTabBar;

};




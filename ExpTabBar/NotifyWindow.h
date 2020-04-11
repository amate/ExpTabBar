#pragma once

// Forward Declare
class CDonutTabBar;

#define	WM_ISMARGECONTROLPANEL	(WM_APP + 1)


////////////////////////////////////////////////////////////
// CNotifyWindow	: 通知用のダミーのウィンドウ

class CNotifyWindow : public CFrameWindowImpl<CNotifyWindow>
{
public:
	DECLARE_FRAME_WND_CLASS(_T("ExpTabBar_NotifyWindow"), NULL)

	// Constructor
	CNotifyWindow(CDonutTabBar* p);

	BEGIN_MSG_MAP_EX(CNotifyWindow)
		MSG_WM_CREATE(OnCreate)
		MSG_WM_DESTROY(OnDestroy)
		MSG_WM_COPYDATA(OnCopyData)
		MESSAGE_HANDLER_EX(WM_ISMARGECONTROLPANEL, OnIsMargeControlPanel)
		//CHAIN_MSG_MAP( CFrameWindowImpl<CNotifyWindow> )
		END_MSG_MAP()

	int OnCreate(LPCREATESTRUCT lpCreateStruct);
	void OnDestroy();
	BOOL OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct);
	LRESULT OnIsMargeControlPanel(UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
	CDonutTabBar* m_pTabBar;

	HANDLE	m_hEventAPIHookTrapper;
	HANDLE	m_hEventAPIHookTrapper64;
	HANDLE	m_hJob;
};



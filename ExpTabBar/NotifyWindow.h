#pragma once

#include <vector>

// Forward Declare
class CDonutTabBar;

#define	WM_ISMARGECONTROLPANEL	(WM_APP + 1)


////////////////////////////////////////////////////////////
// CNotifyWindow	: 通知用のダミーのウィンドウ

class CNotifyWindow : public CFrameWindowImpl<CNotifyWindow>
{
public:
	DECLARE_FRAME_WND_CLASS(_T("ExpTabBar_NotifyWindow"), NULL)

	enum { 
		kFindNextTopActiveTabTimerId = 1, 
	};

	static CNotifyWindow& GetInstance();

	void	AddTabBar(CDonutTabBar* tabBar);
	void	RemoveTabBar(CDonutTabBar* tabBar);

	void	ActiveTabBar(CDonutTabBar* tabBar);

	std::size_t	GetTabBarCount() const {
		return m_tabBarList.size();
	}

	BEGIN_MSG_MAP_EX(CNotifyWindow)
		MSG_WM_CREATE(OnCreate)
		MSG_WM_DESTROY(OnDestroy)
		MSG_WM_COPYDATA(OnCopyData)
		MSG_WM_TIMER(OnTimer)
		MESSAGE_HANDLER_EX(WM_ISMARGECONTROLPANEL, OnIsMargeControlPanel)
		//CHAIN_MSG_MAP( CFrameWindowImpl<CNotifyWindow> )
		END_MSG_MAP()

	int OnCreate(LPCREATESTRUCT lpCreateStruct);
	void OnDestroy();
	BOOL OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct);
	void OnTimer(UINT_PTR nIDEvent);
	LRESULT OnIsMargeControlPanel(UINT uMsg, WPARAM wParam, LPARAM lParam);

private:
	// Constructor
	CNotifyWindow();

	CDonutTabBar* m_pTabBar;
	std::vector<CDonutTabBar*>	m_tabBarList;

	HANDLE	m_hEventAPIHookTrapper;
	HANDLE	m_hEventAPIHookTrapper64;
	HANDLE	m_hJob;
};



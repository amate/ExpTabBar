// DonutDefine.h

#pragma once

#include <WinUser.h>	



#define	WM_DNP_CHANGEFOCUS		(WM_APP + 1)
#define	WM_DNP_CHANGEADDRESS	(WM_APP + 3)
//#define	WM_DNP_CHANGESTATUSTEXT	(WM_APP + 4)
//#define	WM_DNP_CHANGEPROGRESS	(WM_APP + 5)
#define	WM_DNP_SCRIPTERROR		(WM_APP + 6)
#define	WM_DNP_SHOWSCRIPTERROR	(WM_APP + 7)
#define	WM_DNP_SHOWPRIVACYREPORT	(WM_APP + 8)
#define	WM_DNP_SHOWPSECURITYREPORT	(WM_APP + 9)

#define	WM_USER_MAIN_UPDATELEYOUT	(WM_APP + 10)
#define WM_USER_MAIN_SETDELETEVIEW	(WM_APP + 11)

#define	WM_USER_VIEWCREATED			(WM_APP + 12)

// ステータスバー
#define WM_STATUS_SETICON			(WM_USER + 60)
#define WM_STATUS_SETTIPTEXT		(WM_USER + 61)

#define WM_GET_OWNERDRAWMODE		(WM_USER + 95)


#define	SHAREDATANAME	_T("DonutShareData_")

struct SHAREDATA {
	HWND	m_hParWnd;			// 親ウィンドウ(MainFrame)のハンドル
};



// CMainFrame::OnCreate中のAddSimpleReBarBandの順に合わせること
enum E_DonutCtrlID
{
	CmdBarCtrlID = ATL_IDW_BAND_FIRST,
	ToolBarCtrlID,
	AddressBarCtrlID,
	TabBarCtrlID,
};




















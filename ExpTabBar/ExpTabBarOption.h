/**
*	@file	ExpBabBarOption.h
*	@brief	�I�v�V����
*/
#pragma once

#include <atldlgs.h>


// Extended styles
enum EMtb_Ex {
	MTB_EX_DELAYED				= 0x00000004L,
	MTB_EX_ANCHORCOLOR			= 0x00008000L,	// �\��/���\����F��������
	MTB_EX_MOUSEDOWNSELECT		= 0x00400000L,	// �N���b�N�Ń^�u�I��(Drag&Drop�s��)
	// +mod
	MTB_EX_CTRLTAB_MDI			= 0x00800000L,
};


enum TabAddPos {
	RIGHTPOS	= 0,
	LEFTPOS		= 1,
	ACTIVERIGHT	= 2,
	ACTIVELEFT	= 3,
};

enum TabClickCommand {
	COMMANDNONE		= 0,
	TABCLOSE		= 1,
	OPEN_UPFOLDER	= 2,
	NAVIGATELOCK	= 3,
	SHOWMENU		= 4,
};


/////////////////////////////////////////
// �^�u�̐ݒ�

class CTabBarConfig
{
public:
	static int		s_nAddPos;
	static bool		s_bLeftActiveOnClose;
	static bool		s_bAddLinkRight;
	static bool		s_bLinkActive;
	static bool		s_bWheel;
	static bool		s_bMultiLine;
	static bool		s_bUseFixedSize;
	static CSize	s_FixedSize;
	static int		s_nMaxTextLength;
	static int		s_RClickCommand;
	static int		s_DblClickCommand;
	static int		s_MClickCommand;
	static int		s_nMaxHistoryCount;
	
	static void	LoadConfig();
	static void	SaveConfig();
};


///////////////////////////////////////////////////////////
///  �I�v�V�����_�C�A���O

class CExpTabBarOption : public CPropertySheetImpl<CExpTabBarOption>
{
public:
	INT_PTR	Show(HWND hWndParent);

    BEGIN_MSG_MAP_EX( CExpTabBarOption )
        CHAIN_MSG_MAP( CPropertySheetImpl<CExpTabBarOption> )
    END_MSG_MAP()

};















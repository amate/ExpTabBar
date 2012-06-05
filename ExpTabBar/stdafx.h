// stdafx.h : �W���̃V�X�e�� �C���N���[�h �t�@�C���̃C���N���[�h �t�@�C���A�܂���
// �Q�Ɖ񐔂������A�����܂�ύX����Ȃ��A�v���W�F�N�g��p�̃C���N���[�h �t�@�C��
// ���L�q���܂��B

#pragma once

#ifndef STRICT
#define STRICT
#endif
/*
#define WINVER			0x0601 // Windows 7
#define _WIN32_WINNT	0x0601 // Windows 7
#define _WIN32_IE		0x0800 // Internet Explorer 8
*/
#include "targetver.h"

#define _ATL_APARTMENT_THREADED
//#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _WTL_NO_CSTRING		// ATL��CString���g��

//////////////////////////////////
// STL
#include <vector>
#include <algorithm>
#include <thread>
using namespace std::placeholders;

//////////////////////////////////
// boost
#include <boost/tokenizer.hpp>

#include <comsvcs.h>
#include "resource.h"

////////////////////////////////////////////////////////////////////////////////
//ATL���g���ׂ̃t�@�C��

#include <atlstr.h>	// ATL::CString

#include <atlbase.h>
#include <atlapp.h>
#include <atlcom.h>
#include <atlconv.h>
#include <atlwin.h>
#include <atlctl.h>
#include <atlexcept.h>

#include <atlcrack.h>
#include <atlctl.h>

#include <atlframe.h>
#include <atlctrls.h>
#include <atldlgs.h>
#include <atlctrlw.h>

#include <atlctrlx.h>
#include <atlmisc.h>

#include <atlcoll.h>
#include <atltheme.h>


////////////////////////////////////////////////////////////////////////////////
//IDeskBand�Ƃ����g���ׂ̃t�@�C��
#include <shlguid.h>	// IInputObject�Ƃ�������
#include <shlobj.h>		// IDeskBand��IDockingWindow������
#include <tlogstg.h>	// ITravelLogStg�Ƃ�
#include <comdef.h>

#include <xmllite.h>
#pragma comment(lib, "xmllite.lib")


extern HINSTANCE		g_hInst;


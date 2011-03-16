// dllmain.cpp : DllMain の実装

#include "stdafx.h"
#include "resource.h"
#include "ExpTabBar_i.h"
#include "dllmain.h"
#include "DonutFunc.h"

#ifdef _DEBUG
#include <locale>
#endif

CExpTabBarModule _AtlModule;
HINSTANCE		 g_hInst;
TCHAR			 g_szIniFileName[MAX_PATH];	// Ini file name


// DLL エントリ ポイント
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
#ifdef _DEBUG
	_tsetlocale ( LC_ALL, _T("") ); 
#endif
	g_hInst = hInstance;
	_IniFileNameInit(g_szIniFileName, _MAX_PATH);		// iniファイル名を設定

	return _AtlModule.DllMain(dwReason, lpReserved); 
}

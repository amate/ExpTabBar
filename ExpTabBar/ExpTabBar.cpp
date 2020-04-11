// ExpTabBar.cpp : DLL エクスポートの実装です。

//
// メモ: COM+ 1.0 情報:
//      コンポーネントをインストールするには、Microsoft Transaction Explorer をインストールしてください。
//      登録は既定では行われません。 

#include "stdafx.h"
#include "resource.h"
#include "ExpTabBar_i.h"
#include "dllmain.h"

void	InstallRegKey()
{
#if 0
	if (::GetAsyncKeyState(VK_CONTROL) < 0)
		return;

	{// [-HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags]
		CRegKey rk;
		rk.Open(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell"));
		rk.RecurseDeleteKey(_T("Bags"));
	}

	{
		CRegKey rk;
		rk.Open(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced"));
		rk.SetDWORDValue(_T("FullRowSelect"), 0);
	}

	{
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\Bags\\AllFolders\\Shell"), REG_NONE, REG_OPTION_NON_VOLATILE, KEY_WRITE);
		rk.SetStringValue(_T("KnownFolderDerivedFolderType"), _T("{57807898-8C4F-4462-BB63-71042380B109}"));
		rk.SetStringValue(_T("SniffedFolderType"), _T("Generic"));
	}

	{// ;Generic - Folder Template
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\Bags\\AllFolders\\Shell\\{5C4F28B5-F869-4E84-8E60-F11DB97C5CC7}"), REG_NONE, REG_OPTION_NON_VOLATILE, KEY_WRITE);
		rk.SetDWORDValue(_T("Rev")				, 0);
		rk.SetDWORDValue(_T("FFlags")			, 0x43000001);
		rk.SetStringValue(_T("Vid"), _T("{137E7700-3573-11CF-AE69-08002B2E1262}"));
		rk.SetDWORDValue(_T("Mode")				, 4);
		rk.SetDWORDValue(_T("LogicalViewMode")	, 1);
		rk.SetDWORDValue(_T("IconSize")			, 0x10);
		BYTE ColInfo[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xfd,0xdf,0xdf,0xfd,0x10,
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x04,0x00,0x00,0x00,0x18,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,
			0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x10,0x01,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,
			0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0e,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,
			0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x04,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,
			0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0c,0x00,0x00,0x00,0x50,0x00,0x00,0x00
		};
		//rk.SetBinaryValue(_T("ColInfo"), ColInfo, sizeof(ColInfo));
		BYTE Sort[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x30,0xf1,
			0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x01,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("Sort"), Sort, sizeof(Sort));
		rk.SetDWORDValue(_T("GroupView")		, 0);
		rk.SetStringValue(_T("GroupByKey:FMTID"), _T("{00000000-0000-0000-0000-000000000000}"));
		rk.SetDWORDValue(_T("GroupByKey:PID")	, 0);
		rk.SetDWORDValue(_T("GroupByDirection")	, 1);
	}

	{// ;Documents - Folder Template
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\Bags\\AllFolders\\Shell\\{7d49d726-3c21-4f05-99aa-fdc2c9474656}"), REG_NONE, REG_OPTION_NON_VOLATILE, KEY_WRITE);
		rk.SetDWORDValue(_T("Rev")				, 0);
		rk.SetDWORDValue(_T("FFlags")			, 0x43000001);
		rk.SetStringValue(_T("Vid"), _T("{137E7700-3573-11CF-AE69-08002B2E1262}"));
		rk.SetDWORDValue(_T("Mode")				, 4);
		rk.SetDWORDValue(_T("LogicalViewMode")	, 1);
		rk.SetDWORDValue(_T("IconSize")			, 0x10);
		BYTE ColInfo[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xfd,0xdf,0xdf,0xfd,0x10,
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x04,0x00,0x00,0x00,0x18,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,
			0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x10,0x01,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,
			0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0e,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,
			0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x04,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,
			0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0c,0x00,0x00,0x00,0x50,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("ColInfo"), ColInfo, sizeof(ColInfo));
		BYTE Sort[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x30,0xf1,
			0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x01,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("Sort"), Sort, sizeof(Sort));
		rk.SetDWORDValue(_T("GroupView")		, 0);
		rk.SetStringValue(_T("GroupByKey:FMTID"), _T("{00000000-0000-0000-0000-000000000000}"));
		rk.SetDWORDValue(_T("GroupByKey:PID")	, 0);
		rk.SetDWORDValue(_T("GroupByDirection")	, 1);
	}

	{// ;Music - Folder Template
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\Bags\\AllFolders\\Shell\\{94d6ddcc-4a68-4175-a374-bd584a510b78}"), REG_NONE, REG_OPTION_NON_VOLATILE, KEY_WRITE);
		rk.SetDWORDValue(_T("Rev")				, 0);
		rk.SetDWORDValue(_T("FFlags")			, 0x43000001);
		rk.SetStringValue(_T("Vid"), _T("{137E7700-3573-11CF-AE69-08002B2E1262}"));
		rk.SetDWORDValue(_T("Mode")				, 4);
		rk.SetDWORDValue(_T("LogicalViewMode")	, 1);
		rk.SetDWORDValue(_T("IconSize")			, 0x10);
		BYTE ColInfo[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xfd,0xdf,0xdf,0xfd,0x10,
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x04,0x00,0x00,0x00,0x18,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,
			0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x10,0x01,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,
			0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0e,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,
			0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x04,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,
			0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0c,0x00,0x00,0x00,0x50,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("ColInfo"), ColInfo, sizeof(ColInfo));
		BYTE Sort[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x30,0xf1,
			0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x01,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("Sort"), Sort, sizeof(Sort));
		rk.SetDWORDValue(_T("GroupView")		, 0);
		rk.SetStringValue(_T("GroupByKey:FMTID"), _T("{00000000-0000-0000-0000-000000000000}"));
		rk.SetDWORDValue(_T("GroupByKey:PID")	, 0);
		rk.SetDWORDValue(_T("GroupByDirection")	, 1);
	}

	{// ;Pictures - Folder Template
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\Bags\\AllFolders\\Shell\\{b3690e58-e961-423b-b687-386ebfd83239}"), REG_NONE, REG_OPTION_NON_VOLATILE, KEY_WRITE);
		rk.SetDWORDValue(_T("Rev")				, 0);
		rk.SetDWORDValue(_T("FFlags")			, 0x43000001);
		rk.SetStringValue(_T("Vid"), _T("{137E7700-3573-11CF-AE69-08002B2E1262}"));
		rk.SetDWORDValue(_T("Mode")				, 4);
		rk.SetDWORDValue(_T("LogicalViewMode")	, 1);
		rk.SetDWORDValue(_T("IconSize")			, 0x10);
		BYTE ColInfo[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xfd,0xdf,0xdf,0xfd,0x10,
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x04,0x00,0x00,0x00,0x18,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,
			0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x10,0x01,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,
			0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0e,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,
			0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x04,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,
			0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0c,0x00,0x00,0x00,0x50,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("ColInfo"), ColInfo, sizeof(ColInfo));
		BYTE Sort[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x30,0xf1,
			0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x01,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("Sort"), Sort, sizeof(Sort));
		rk.SetDWORDValue(_T("GroupView")		, 0);
		rk.SetStringValue(_T("GroupByKey:FMTID"), _T("{00000000-0000-0000-0000-000000000000}"));
		rk.SetDWORDValue(_T("GroupByKey:PID")	, 0);
		rk.SetDWORDValue(_T("GroupByDirection")	, 1);
	}

	{// ;Videos - Folder Template
		CRegKey rk;
		rk.Create(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell\\Bags\\AllFolders\\Shell\\{5fa96407-7e77-483c-ac93-691d05850de8}"), REG_NONE, REG_OPTION_NON_VOLATILE, KEY_WRITE);
		rk.SetDWORDValue(_T("Rev")				, 0);
		rk.SetDWORDValue(_T("FFlags")			, 0x43000001);
		rk.SetStringValue(_T("Vid"), _T("{137E7700-3573-11CF-AE69-08002B2E1262}"));
		rk.SetDWORDValue(_T("Mode")				, 4);
		rk.SetDWORDValue(_T("LogicalViewMode")	, 1);
		rk.SetDWORDValue(_T("IconSize")			, 0x10);
		BYTE ColInfo[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0xfd,0xdf,0xdf,0xfd,0x10,
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x04,0x00,0x00,0x00,0x18,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,
			0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x10,0x01,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,
			0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0e,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,
			0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x04,0x00,0x00,0x00,0x78,0x00,0x00,0x00,0x30,0xf1,0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,
			0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0c,0x00,0x00,0x00,0x50,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("ColInfo"), ColInfo, sizeof(ColInfo));
		BYTE Sort[] = { 
			0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x30,0xf1,
			0x25,0xb7,0xef,0x47,0x1a,0x10,0xa5,0xf1,0x02,0x60,0x8c,0x9e,0xeb,0xac,0x0a,0x00,0x00,0x00,0x01,0x00,0x00,0x00
		};
		rk.SetBinaryValue(_T("Sort"), Sort, sizeof(Sort));
		rk.SetDWORDValue(_T("GroupView")		, 0);
		rk.SetStringValue(_T("GroupByKey:FMTID"), _T("{00000000-0000-0000-0000-000000000000}"));
		rk.SetDWORDValue(_T("GroupByKey:PID")	, 0);
		rk.SetDWORDValue(_T("GroupByDirection")	, 1);
	}
#endif
}

void	UnInstallRegKey()
{
#if 0
	{// [-HKEY_CURRENT_USER\Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\Bags]
		CRegKey rk;
		rk.Open(HKEY_CURRENT_USER, _T("Software\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\Shell"));
		rk.RecurseDeleteKey(_T("Bags"));
	}

	{
		CRegKey rk;
		rk.Open(HKEY_CURRENT_USER, _T("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced"));
		rk.DeleteValue(_T("FullRowSelect"));
	}
#endif
}

// Used to determine whether the DLL can be unloaded by OLE.
STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}

// Returns a class factory to create an object of the requested type.
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

// DllRegisterServer - Adds entries to the system registry.
STDAPI DllRegisterServer(void)
{
	InstallRegKey();

	// オブジェクト、タイプ ライブラリおよびタイプ ライブラリ内のすべてのインターフェイスを登録します
	HRESULT hr = _AtlModule.DllRegisterServer();
	return hr;
}

// DllUnregisterServer - Removes entries from the system registry.
STDAPI DllUnregisterServer(void)
{
	UnInstallRegKey();

	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

// DllInstall - Adds/Removes entries to the system registry per user per machine.
STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
	HRESULT hr = E_FAIL;
	static const wchar_t szUserSwitch[] = L"user";

	if (pszCmdLine != NULL)
	{
		if (_wcsnicmp(pszCmdLine, szUserSwitch, _countof(szUserSwitch)) == 0)
		{
			ATL::AtlSetPerUserRegistration(true);
		}
	}

	if (bInstall)
	{	
		hr = DllRegisterServer();
		if (FAILED(hr))
		{
			DllUnregisterServer();
		}
	}
	else
	{
		hr = DllUnregisterServer();
	}

	return hr;
}



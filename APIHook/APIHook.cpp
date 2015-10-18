// APIHook.cpp : DLL アプリケーション用にエクスポートされる関数を定義します。
//

#include "stdafx.h"
#include <assert.h>
#include <shellapi.h>
#include <ShlObj.h>
#include <Shlwapi.h>
#pragma comment(lib, "Shlwapi.lib")
#include <psapi.h>
#pragma comment(lib, "psapi.lib")
#include <boost\format.hpp>
#include <regex>

#define EXPORTS
#include "APIHook.h"
#include "..\minhook\include\MinHook.h"
#include "OpenFolderAndSelectItems.h"
#include "Logger.h"

extern HMODULE g_hInst;

#pragma data_seg(".sharedata")
HHOOK g_hHook = NULL;
#pragma data_seg()
#pragma comment(linker, "/section:.sharedata,rws")

bool	g_HookInstall = false;

void	ExpTabBarFolderOpen(HWND hWndTarget, LPITEMIDLIST pidl)
{
	assert(pidl);

	COPYDATASTRUCT	cd = { 0 };
	UINT cbItemID = ::ILGetSize(pidl);
	cd.lpData = (LPVOID)pidl;
	cd.cbData = cbItemID;
	SendMessage(hWndTarget, WM_COPYDATA, NULL, (LPARAM)&cd);
	::CoTaskMemFree(pidl);
}

void	ExpTabBarFolderOpen(HWND hWndTarget, const OpenFolderAndSelectItems& folderAndSelectItems)
{
	std::string serializedData = folderAndSelectItems.Serialize();

	COPYDATASTRUCT	cd = { 0 };
	cd.dwData = 1;
	cd.lpData = (LPVOID)serializedData.data();
	cd.cbData = (DWORD)serializedData.length();
	SendMessage(hWndTarget, WM_COPYDATA, NULL, (LPARAM)&cd);
}

bool GetWindowFileName(HWND hWnd, LPTSTR lpFileName, DWORD nSize)
{
	DWORD processID;
	GetWindowThreadProcessId(hWnd, &processID);
	HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processID);
	bool ret = false;
	if (hProcess)
	{
		HMODULE hModule;
		DWORD cbReturned;
		if (EnumProcessModules(hProcess, &hModule, sizeof(hModule), &cbReturned))
		{
			ret = 0 != GetModuleFileNameEx(hProcess, hModule, lpFileName, nSize);
		}
		CloseHandle(hProcess);
	}

	return ret;
}


// ShellExecuteA Hook
using func_ShellExecuteA = HINSTANCE (WINAPI *)(_In_opt_ HWND hwnd, _In_opt_ LPCSTR lpOperation, _In_ LPCSTR lpFile, _In_opt_ LPCSTR lpParameters,
	_In_opt_ LPCSTR lpDirectory, _In_ INT nShowCmd);
func_ShellExecuteA orgfunc_ShellExecuteA;

HINSTANCE WINAPI ShellExecuteA_Hook(_In_opt_ HWND hwnd, _In_opt_ LPCSTR lpOperation, _In_ LPCSTR lpFile, _In_opt_ LPCSTR lpParameters,
	_In_opt_ LPCSTR lpDirectory, _In_ INT nShowCmd)
{
	if (::PathIsDirectoryA(lpFile) &&
		(lpOperation == nullptr || ::lstrcmpA(lpOperation, "open") == 0 || ::lstrcmpA(lpOperation, "explore") == 0))
	{
		HWND hWndTarget = FindWindow(L"ExpTabBar_NotifyWindow", NULL);
		if (hWndTarget) {
			// 他のエクスプローラーにpidlが開かれたことを通知する
			LPITEMIDLIST pidl = ILCreateFromPathA(lpFile);
			if (pidl) {
				ExpTabBarFolderOpen(hWndTarget, pidl);
				return (HINSTANCE)32;
			}
		}
	}

	//std::string text = (boost::format("lpOperation : %1% lpFile : %2% lpParameters : %3%") % lpOperation % lpFile % lpParameters).str().c_str();
	//MessageBoxA(NULL, text.c_str(), NULL, MB_OK);

	return orgfunc_ShellExecuteA(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
}

// ShellExecuteW Hook
using func_ShellExecuteW = HINSTANCE(WINAPI *)(_In_opt_ HWND hwnd, _In_opt_ LPCWSTR lpOperation, _In_ LPCWSTR lpFile, _In_opt_ LPCWSTR lpParameters,
	_In_opt_ LPCWSTR lpDirectory, _In_ INT nShowCmd);
func_ShellExecuteW orgfunc_ShellExecuteW;

HINSTANCE WINAPI ShellExecuteW_Hook(_In_opt_ HWND hwnd, _In_opt_ LPCWSTR lpOperation, _In_ LPCWSTR lpFile, _In_opt_ LPCWSTR lpParameters,
	_In_opt_ LPCWSTR lpDirectory, _In_ INT nShowCmd)
{
	if (::PathIsDirectoryW(lpFile) && 
		(lpOperation == nullptr || ::lstrcmpW(lpOperation, L"open") == 0 || ::lstrcmpW(lpOperation, L"explore") == 0) )
	{
		HWND hWndTarget = FindWindow(L"ExpTabBar_NotifyWindow", NULL);
		if (hWndTarget) {
			// 他のエクスプローラーにpidlが開かれたことを通知する
			LPITEMIDLIST pidl = ILCreateFromPathW(lpFile);
			if (pidl) {
				ExpTabBarFolderOpen(hWndTarget, pidl);
				return (HINSTANCE)32;
			}
		}
	}
	//std::wstring text = (boost::wformat(L"lpOperation : %1% lpFile : %2% lpParameters : %3%") % lpOperation % lpFile % lpParameters).str().c_str();
	//MessageBoxW(NULL, text.c_str(), NULL, MB_OK);

	return orgfunc_ShellExecuteW(hwnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd);
}

// SHOpenFolderAndSelectItems Hook
using func_SHOpenFolderAndSelectItems = HRESULT (WINAPI *)(_In_ PCIDLIST_ABSOLUTE pidlFolder, UINT cidl, _In_reads_opt_(cidl) PCUITEMID_CHILD_ARRAY apidl, DWORD dwFlags);
func_SHOpenFolderAndSelectItems orgfunc_SHOpenFolderAndSelectItems;

HRESULT WINAPI SHOpenFolderAndSelectItems_Hook(_In_ PCIDLIST_ABSOLUTE pidlFolder, UINT cidl, _In_reads_opt_(cidl) PCUITEMID_CHILD_ARRAY apidl, DWORD dwFlags)
{
	HWND hWndTarget = FindWindow(L"ExpTabBar_NotifyWindow", NULL);
	if (hWndTarget) {
		OpenFolderAndSelectItems folderAndSelectItems;
		std::string folder((const char*)pidlFolder, ::ILGetSize(pidlFolder));
		folderAndSelectItems.pidlFolder = folder;
		folderAndSelectItems.cidl = cidl;
		for (UINT i = 0; i < cidl; ++i) {
			std::string childItem((const char*)apidl[i], ::ILGetSize(apidl[i]));
			folderAndSelectItems.apidl.emplace_back(childItem);
		}
		folderAndSelectItems.dwFlags = dwFlags;

		ExpTabBarFolderOpen(hWndTarget, folderAndSelectItems);

		return S_OK;
	}

	return orgfunc_SHOpenFolderAndSelectItems(pidlFolder, cidl, apidl, dwFlags);
}

// CreateProcessA Hook
using func_CreateProcessA = BOOL (WINAPI *)(_In_opt_ LPCSTR lpApplicationName,
											_Inout_opt_ LPSTR lpCommandLine,
											_In_opt_ LPSECURITY_ATTRIBUTES lpProcessAttributes,
											_In_opt_ LPSECURITY_ATTRIBUTES lpThreadAttributes,
											_In_ BOOL bInheritHandles,
											_In_ DWORD dwCreationFlags,
											_In_opt_ LPVOID lpEnvironment,
											_In_opt_ LPCSTR lpCurrentDirectory,
											_In_ LPSTARTUPINFOA lpStartupInfo,
											_Out_ LPPROCESS_INFORMATION lpProcessInformation
											);
func_CreateProcessA orgfunc_CreateProcessA;

BOOL WINAPI CreateProcessA_Hook(_In_opt_ LPCSTR lpApplicationName,
	_Inout_opt_ LPSTR lpCommandLine,
	_In_opt_ LPSECURITY_ATTRIBUTES lpProcessAttributes,
	_In_opt_ LPSECURITY_ATTRIBUTES lpThreadAttributes,
	_In_ BOOL bInheritHandles,
	_In_ DWORD dwCreationFlags,
	_In_opt_ LPVOID lpEnvironment,
	_In_opt_ LPCSTR lpCurrentDirectory,
	_In_ LPSTARTUPINFOA lpStartupInfo,
	_Out_ LPPROCESS_INFORMATION lpProcessInformation
	)
{
	//INFO_LOG << L"CreateProcessA_Hook AppName : " << lpApplicationName << L" CmdLine : " << lpCommandLine << L" ProcesInfo : " << lpProcessInformation;
	if (::lstrcmpiA(lpApplicationName, R"(C:\Windows\explorer.exe)") == 0) {
		HWND hWndTarget = FindWindow(L"ExpTabBar_NotifyWindow", NULL);
		if (hWndTarget) {
			std::regex rx(R"*(/select,"?([^"]+)"?)*");
			std::smatch result;
			std::string cmdLine = lpCommandLine;
			if (std::regex_search(cmdLine, result, rx)) {
				std::string path = result[1].str();
				if (::PathIsDirectoryA(path.c_str()) || ::PathFileExistsA(path.c_str())) {
					OpenFolderAndSelectItems folderAndSelectItems;
					LPITEMIDLIST pidl = ::ILCreateFromPathA(path.c_str());
					std::string folder((const char*)pidl, ::ILGetSize(pidl));
					::ILFree(pidl);
					folderAndSelectItems.pidlFolder = folder;

					//INFO_LOG << L"ExpTabBarFolderOpen : " << path;
					ExpTabBarFolderOpen(hWndTarget, folderAndSelectItems);
					return TRUE;
				}
			}
		}
	}
	return orgfunc_CreateProcessA(lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
}

// CreateProcessW Hook
using func_CreateProcessW = BOOL (WINAPI *)(_In_opt_ LPCWSTR lpApplicationName,
											_Inout_opt_ LPWSTR lpCommandLine,
											_In_opt_ LPSECURITY_ATTRIBUTES lpProcessAttributes,
											_In_opt_ LPSECURITY_ATTRIBUTES lpThreadAttributes,
											_In_ BOOL bInheritHandles,
											_In_ DWORD dwCreationFlags,
											_In_opt_ LPVOID lpEnvironment,
											_In_opt_ LPCWSTR lpCurrentDirectory,
											_In_ LPSTARTUPINFOW lpStartupInfo,
											_Out_ LPPROCESS_INFORMATION lpProcessInformation
											);
func_CreateProcessW orgfunc_CreateProcessW;

BOOL WINAPI CreateProcessW_Hook(
	_In_opt_ LPCWSTR lpApplicationName,
	_Inout_opt_ LPWSTR lpCommandLine,
	_In_opt_ LPSECURITY_ATTRIBUTES lpProcessAttributes,
	_In_opt_ LPSECURITY_ATTRIBUTES lpThreadAttributes,
	_In_ BOOL bInheritHandles,
	_In_ DWORD dwCreationFlags,
	_In_opt_ LPVOID lpEnvironment,
	_In_opt_ LPCWSTR lpCurrentDirectory,
	_In_ LPSTARTUPINFOW lpStartupInfo,
	_Out_ LPPROCESS_INFORMATION lpProcessInformation
	)
{
	//INFO_LOG << L"CreateProcessW_Hook AppName : " << lpApplicationName << L" CmdLine : " << lpCommandLine << L" ProcesInfo : " << lpProcessInformation;
	if (::lstrcmpiW(lpApplicationName, LR"(C:\Windows\explorer.exe)") == 0) {
		HWND hWndTarget = FindWindow(L"ExpTabBar_NotifyWindow", NULL);
		if (hWndTarget) {
			std::wregex rx(LR"*(/select,"?([^"]+)"?)*");
			std::wsmatch result;
			std::wstring cmdLine = lpCommandLine;
			if (std::regex_search(cmdLine, result, rx)) {
				std::wstring path = result[1].str();
				if (::PathIsDirectoryW(path.c_str()) || ::PathFileExistsW(path.c_str())) {
					OpenFolderAndSelectItems folderAndSelectItems;
					LPITEMIDLIST pidl = ::ILCreateFromPathW(path.c_str());
					std::string folder((const char*)pidl, ::ILGetSize(pidl));
					::ILFree(pidl);
					folderAndSelectItems.pidlFolder = folder;

					//INFO_LOG << L"ExpTabBarFolderOpen : " << path;
					ExpTabBarFolderOpen(hWndTarget, folderAndSelectItems);
					return TRUE;
				}
			}
		}
	}
	return orgfunc_CreateProcessW(lpApplicationName, lpCommandLine, lpProcessAttributes, lpThreadAttributes, bInheritHandles, dwCreationFlags, lpEnvironment, lpCurrentDirectory, lpStartupInfo, lpProcessInformation);
}


class CHookUnHookManager
{
public:
	CHookUnHookManager() : m_bHookSuccess(false)
	{
		if (MH_Initialize() == MH_OK) {
			m_bHookSuccess = true;
			MH_STATUS ret = MH_OK;
			ret = MH_CreateHookApi(L"shell32", "ShellExecuteA", (LPVOID)&ShellExecuteA_Hook, (LPVOID*)&orgfunc_ShellExecuteA);
			assert(ret == MH_OK);

			ret = MH_CreateHookApi(L"shell32", "ShellExecuteW", (LPVOID)&ShellExecuteW_Hook, (LPVOID*)&orgfunc_ShellExecuteW);
			assert(ret == MH_OK);

			ret = MH_CreateHookApi(L"shell32", "SHOpenFolderAndSelectItems", (LPVOID)&SHOpenFolderAndSelectItems_Hook, (LPVOID*)&orgfunc_SHOpenFolderAndSelectItems);
			assert(ret == MH_OK);

			ret = MH_CreateHookApi(L"Kernel32", "CreateProcessA", (LPVOID)&CreateProcessA_Hook, (LPVOID*)&orgfunc_CreateProcessA);
			assert(ret == MH_OK);

			ret = MH_CreateHookApi(L"Kernel32", "CreateProcessW", (LPVOID)&CreateProcessW_Hook, (LPVOID*)&orgfunc_CreateProcessW);
			assert(ret == MH_OK);
			
			ret = MH_EnableHook(MH_ALL_HOOKS);
			assert(ret == MH_OK);
		}
	}

	~CHookUnHookManager()
	{
		if (m_bHookSuccess) {
			MH_Uninitialize();
		}
	}

private:
	bool m_bHookSuccess;

};

void	InstallHook()
{
	static CHookUnHookManager hookManager;
}

LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam)
{
	if (nCode == HCBT_ACTIVATE) {
		if (g_HookInstall == false) {
			g_HookInstall = true;

			// エクスプローラーはフックしない
			HWND hWndActivate = (HWND)wParam;
			TCHAR exeFilePath[MAX_PATH];
			if (GetWindowFileName(hWndActivate, exeFilePath, MAX_PATH)) {
				LPCTSTR fileName = ::PathFindFileName(exeFilePath);
				if (fileName && ::lstrcmpi(fileName, L"explorer.exe") == 0) {
					goto SkipInstall;
				}
			}
			InstallHook();

			SkipInstall:;
		}
	}
	return CallNextHookEx(g_hHook, nCode, wParam, lParam);
}


int InitHook()
{
	g_hHook = ::SetWindowsHookEx(WH_CBT, CBTProc, g_hInst, 0);
	if (g_hHook == NULL)
		return -1;

	return 1;
}

int TermHook()
{
	if (g_hHook == NULL)
		return -1;

	::UnhookWindowsHookEx(g_hHook);
	g_hHook = NULL;

	return 1;
}
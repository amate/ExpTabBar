#include "stdafx.h"
#include "NotifyWindow.h"
#include "DonutTabBar.h"
#include "ShellWrap.h"
#include "ExpTabBarOption.h"


////////////////////////////////////////////////////////////
// CNotifyWindow

// Constructor
CNotifyWindow::CNotifyWindow(CDonutTabBar* p)
	: m_pTabBar(p), m_hEventAPIHookTrapper(NULL), m_hEventAPIHookTrapper64(NULL), m_hJob(NULL)
{	}

int CNotifyWindow::OnCreate(LPCREATESTRUCT lpCreateStruct)
{
	if (CTabBarConfig::s_bUseAPIHook) {
		LPCWSTR kJobName = L"ExpTabBar_Job";
		m_hJob = ::CreateJobObject(nullptr, kJobName);
		ATLASSERT(m_hJob);
		JOBOBJECT_EXTENDED_LIMIT_INFORMATION extendedLimit = {};
		extendedLimit.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
		SetInformationJobObject(m_hJob, JobObjectExtendedLimitInformation, &extendedLimit, sizeof(JOBOBJECT_EXTENDED_LIMIT_INFORMATION));

		auto funcStartProcess = [](const CString& exePath, HANDLE hEvent) {
			if (::PathFileExists(exePath) == FALSE)
				return;

			STARTUPINFO startUpInfo = { sizeof(STARTUPINFO) };
			PROCESS_INFORMATION processInfo = {};
			SECURITY_ATTRIBUTES securityAttributes = { sizeof(SECURITY_ATTRIBUTES) };
			securityAttributes.bInheritHandle = TRUE;
			std::wstring cmdLine = std::to_wstring((uint64_t)hEvent);
			BOOL bRet = ::CreateProcess(exePath, (LPWSTR)cmdLine.data(),
				nullptr, nullptr, TRUE, 0, nullptr, nullptr, &startUpInfo, &processInfo);
			ATLASSERT(bRet);
			::CloseHandle(processInfo.hThread);
			::CloseHandle(processInfo.hProcess);
		};

		SECURITY_ATTRIBUTES securityAttributes = { sizeof(SECURITY_ATTRIBUTES) };
		securityAttributes.bInheritHandle = TRUE;
		m_hEventAPIHookTrapper = ::CreateEvent(&securityAttributes, TRUE, FALSE, NULL);
		ATLASSERT(m_hEventAPIHookTrapper);
		funcStartProcess(Misc::GetExeDirectory() + L"APIHookTrapper.exe", m_hEventAPIHookTrapper);

#ifdef _WIN64
		m_hEventAPIHookTrapper64 = ::CreateEvent(&securityAttributes, TRUE, FALSE, NULL);
		ATLASSERT(m_hEventAPIHookTrapper64);
		funcStartProcess(Misc::GetExeDirectory() + L"APIHookTrapper64.exe", m_hEventAPIHookTrapper64);
#endif

	}
	return 0;
}

void CNotifyWindow::OnDestroy()
{
	if (m_hJob) {
		::SetEvent(m_hEventAPIHookTrapper);
		::Sleep(1);
		::CloseHandle(m_hEventAPIHookTrapper);
		m_hEventAPIHookTrapper = NULL;

#ifdef _WIN64
		::SetEvent(m_hEventAPIHookTrapper64);
		::Sleep(1);
		::CloseHandle(m_hEventAPIHookTrapper64);
		m_hEventAPIHookTrapper64 = NULL;
#endif

		::CloseHandle(m_hJob);
		m_hJob = NULL;
	}
}

/// 外部からタブを開く
BOOL CNotifyWindow::OnCopyData(CWindow wnd, PCOPYDATASTRUCT pCopyDataStruct)
{
#if 0
	LPCTSTR strFullPath = (LPCTSTR)pCopyDataStruct->lpData;
#endif
	if (pCopyDataStruct->dwData == 0) {
		LPITEMIDLIST pidl = (LPITEMIDLIST)pCopyDataStruct->lpData;
		if (pidl == NULL)
			return FALSE;

		CString folderPath = GetFullPathFromIDList(pidl);
		//INFO_LOG << L"ExternalOpen 0: " << (LPCWSTR)folderPath;

		m_pTabBar->ExternalOpen(::ILClone(pidl));

	} else if (pCopyDataStruct->dwData == 1) {
		std::string serializedData((const char*)pCopyDataStruct->lpData, pCopyDataStruct->cbData);
		OpenFolderAndSelectItems folderAndSelectItems;
		folderAndSelectItems.Deserialize(serializedData);

#if 0
		// 子がフルパスになっていたら、子だけにする
		for (UINT i = 0; i < folderAndSelectItems.cidl; ++i) {
			LPITEMIDLIST pidl = (LPITEMIDLIST)folderAndSelectItems.apidl[i].data();
			bool isChild = ILIsChild(pidl) != 0;
			if (isChild == false) {
				PITEMID_CHILD childpidl = ::ILFindLastID(pidl);
				if (childpidl == nullptr) {
					ATLASSERT(FALSE);
				} else {
					std::string childItem((const char*)childpidl, ::ILGetSize(childpidl));
					folderAndSelectItems.apidl[i] = childItem;
				}
			}
		}
#endif
		// pidlFolderがフォルダでなければ、分けて入れる
		LPCITEMIDLIST pidl = (LPCITEMIDLIST)folderAndSelectItems.pidlFolder.data();
		if (IsExistFolderFromIDList(pidl) == false) {
			ATLASSERT(folderAndSelectItems.cidl == 0);
			PITEMID_CHILD childpidl = ::ILFindLastID(pidl);
			if (childpidl == nullptr) {
				ATLASSERT(FALSE);
			} else {
				std::string childItem((const char*)childpidl, ::ILGetSize(childpidl));
				folderAndSelectItems.apidl.emplace_back(childItem);
				++folderAndSelectItems.cidl;

				LPITEMIDLIST pidlClone = ::ILClone(pidl);
				::ILRemoveLastID(pidlClone);
				std::string folderpidl((const char*)pidlClone, ::ILGetSize(pidlClone));
				folderAndSelectItems.pidlFolder = folderpidl;
				::CoTaskMemFree(pidlClone);
			}
		}

		CString folderPath = GetFullPathFromIDList((LPCITEMIDLIST)folderAndSelectItems.pidlFolder.data());
		//INFO_LOG << L"ExternalOpen 1: " << (LPCWSTR)folderPath << L" cidl : " << folderAndSelectItems.cidl << L" dwFlags : " << folderAndSelectItems.dwFlags;

		m_pTabBar->ExternalOpen(::ILClone((LPCITEMIDLIST)folderAndSelectItems.pidlFolder.data()), folderAndSelectItems);
	}
	return TRUE;
}

LRESULT CNotifyWindow::OnIsMargeControlPanel(UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	return CTabBarConfig::s_bMargeControlPanel;
}


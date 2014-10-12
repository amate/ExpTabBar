// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved
//
// This demonstrates how implement a shell verb using the ExecuteCommand method
// this method is preferred for verb implementations as it provides the most flexibility,
// it is simple, and supports out of process activation.
//
// This sample implements a standalone local server COM object but
// it is expected that the verb implementation will be integrated
// into existing applications. to do that have your main application object
// register a class factory for itself and have that object implement IDropTarget
// for the verbs of your application. Note that COM will launch your application
// if it is not already running and will connect to an already running instance
// of your application if it is already running. These are features of the COM
// based verb implementation methods.
//
// It is also possible (but not recommended) to create in process implementations
// of this object. To do that follow this sample but replace the local
// server COM object with an in-proc server.

/**
*	@file	openInTab.cpp
*	@brief	ExpTabBarに開くタブの情報を渡す
*/

#include "ShellHelpers.h"
#include "VerbHelpers.h"
#include "RegisterExtension.h"
#include <strsafe.h>
#include <new>  // std::nothrow
#include <assert.h>

// Each ExecuteCommand handler needs to have a unique COM object, run UUIDGEN.EXE to
// create new CLSID values for your handler. These handlers can implement multiple
// different verbs using the information provided via IInitializeCommand, for example the verb name.
// your code can switch off those different verb names or the properties provided
// in the property bag.

WCHAR const c_szVerbDisplayName[] = L"タブで開く2";

#define EXPTABBAR_NOTIFYWINDOWCLASSNAME	L"ExpTabBar_NotifyWindow"

#define	WM_ISMARGECONTROLPANEL	(WM_APP + 1)

void	UnRegisterExecuteCommandVerb()
{
	HKEY hkey;
	if (RegOpenKeyEx(HKEY_CURRENT_USER, L"Software\\Classes\\Folder\\shell", 0, KEY_ALL_ACCESS, &hkey) == ERROR_SUCCESS) {
		RegDeleteTree(hkey, L"open");
		RegCloseKey(hkey);
	}
}

class __declspec(uuid("597EEB5A-D2DD-46E1-8BD2-C03CF13B8C3E"))
    CExecuteCommandVerb : public IExecuteCommand,
                          public IObjectWithSelection,
                          public IInitializeCommand,
                          public IObjectWithSite,
                          CAppMessageLoop
{
public:
	CExecuteCommandVerb() : _cRef(1), _psia(NULL), _punkSite(NULL)
    {
    }

    HRESULT Run();

    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv)
    {		
        static const QITAB qit[] =
        {
            QITABENT(CExecuteCommandVerb, IExecuteCommand),        // required
            QITABENT(CExecuteCommandVerb, IObjectWithSelection),   // required
            QITABENT(CExecuteCommandVerb, IInitializeCommand),     // optional
            QITABENT(CExecuteCommandVerb, IObjectWithSite),        // optional
            { 0 },
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef()
    {
        return InterlockedIncrement(&_cRef);
    }

    IFACEMETHODIMP_(ULONG) Release()
    {
        long cRef = InterlockedDecrement(&_cRef);
        if (!cRef)
        {
            delete this;
        }
        return cRef;
    }

    // IExecuteCommand
    IFACEMETHODIMP SetKeyState(DWORD grfKeyState)
    {
        _grfKeyState = grfKeyState;
        return S_OK;
    }

    IFACEMETHODIMP SetParameters(PCWSTR pszParameters)
	{ return S_OK; }

    IFACEMETHODIMP SetPosition(POINT pt)
    {
        _pt = pt;
        return S_OK;
    }

    IFACEMETHODIMP SetShowWindow(int nShow)
    {
        _nShow = nShow;
        return S_OK;
    }

    IFACEMETHODIMP SetNoShowUI(BOOL /* fNoShowUI */)
    { return S_OK; }

    IFACEMETHODIMP SetDirectory(PCWSTR pszDirectory)
    { return S_OK; }

    IFACEMETHODIMP Execute();

    // IObjectWithSelection
    IFACEMETHODIMP SetSelection(IShellItemArray *psia)
    {
        SetInterface(&_psia, psia);
        return S_OK;
    }

    IFACEMETHODIMP GetSelection(REFIID riid, void **ppv)
    {
        *ppv = NULL;
        return _psia ? _psia->QueryInterface(riid, ppv) : E_FAIL;
    }

    // IInitializeCommand
    IFACEMETHODIMP Initialize(PCWSTR pszCommandName, IPropertyBag*  ppb )
	{
        // The verb name is in pszCommandName, this handler can varry its behavior
        // based on the command name (implementing different verbs) or the
        // data stored under that verb in the registry can be read via ppb
        return S_OK;
    }

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown *punkSite)
    {
        SetInterface(&_punkSite, punkSite);
        return S_OK;
    }

    IFACEMETHODIMP GetSite(REFIID riid, void **ppv)
    {
        *ppv = NULL;
        return _punkSite ? _punkSite->QueryInterface(riid, ppv) : E_FAIL;
    }

	void RunExplorer(LPCTSTR strPath)
	{
		LPWSTR strWinFolder = nullptr;
		HRESULT hr = SHGetKnownFolderPath(FOLDERID_Windows, 0, NULL, &strWinFolder);
		if (SUCCEEDED(hr)) {
			WCHAR strExplorerPath[MAX_PATH];
			StringCchPrintf(strExplorerPath, MAX_PATH, L"%s\\explorer.exe", strWinFolder);
			::CoTaskMemFree(strWinFolder);
			::ShellExecute(NULL, NULL, strExplorerPath, strPath, NULL, SW_NORMAL);
		}
	}

	/// ExpTabBarに開かれるフォルダの情報を送る
    void OnAppCallback()
    {
		//assert(FALSE);
		if (_psia == nullptr) {		
			// タスクバーの時計の右クリックメニューのプロパティをクリックしたときに
			// 自前で"通知領域アイコン\\システム アイコン"を開かせる
			//RunExplorer(L"コントロール パネル\\すべてのコントロール パネル項目\\通知領域アイコン\\システム アイコン");
			UnRegisterExecuteCommandVerb();
			MessageBox(NULL, L"一時的にタブ取り込みを無効化しました\nもう一度実行してください", L"openInTabの起動に失敗。", MB_ICONERROR);
			return ;
		}

		HWND hWndNotify = ::FindWindow(EXPTABBAR_NOTIFYWINDOWCLASSNAME, NULL);
		DWORD count = 0;
		_psia->GetCount(&count);
		for (DWORD i = 0; i < count; ++i) {
			IShellItem2 *psi;
			HRESULT hr = GetItemAt(_psia, i, IID_PPV_ARGS(&psi));
			if (SUCCEEDED(hr)) {
				PIDLIST_ABSOLUTE pidl = nullptr;
				hr = ::SHGetIDListFromObject(psi, &pidl);
				if (SUCCEEDED(hr) && pidl) {
					LPWSTR strPath = nullptr;
					hr = psi->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &strPath);

					if (hWndNotify) {
						if (::SendMessage(hWndNotify, WM_ISMARGECONTROLPANEL, 0, 0) != 0 ||		// コントローラーパネルを取り込む
							wcsncmp(strPath, L"::{26EE0668-A00A-44D7-9371-BEB064C98683}", 40) != 0 ) {
							COPYDATASTRUCT	cd = { 0 };
							UINT cbItemID = ::ILGetSize(pidl);
							cd.lpData	= (LPVOID)pidl;
							cd.cbData	= cbItemID;
							SendMessage(hWndNotify, WM_COPYDATA, NULL, (LPARAM)&cd);	
							
						} else {
							RunExplorer(strPath);
						}

					} else {	// ExpTabBarが見つからない
						UnRegisterExecuteCommandVerb();

						RunExplorer(strPath);
					}
					::CoTaskMemFree(strPath);
					::CoTaskMemFree(pidl);
				}
#if 0
				LPWSTR strPath = nullptr;
				hr = psi->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &strPath);
				if (SUCCEEDED(hr)) {
					PIDLIST_ABSOLUTE pidl = nullptr;
					DWORD dwAttr = 0;
					pidl = ::ILCreateFromPath(strPath);
					if (pidl) {
						if (hWndNotify) {
							COPYDATASTRUCT	cd = { 0 };
							UINT cbItemID = ::ILGetSize(pidl);
							cd.lpData	= (LPVOID)pidl;
							cd.cbData	= cbItemID;
							SendMessage(hWndNotify, WM_COPYDATA, NULL, (LPARAM)&cd);
							::CoTaskMemFree(pidl);

						} else {	// ExpTabBarが見つからない
							UnRegisterExecuteCommandVerb();
							::ShellExecute(NULL, NULL, L"explorer", strPath, NULL, SW_NORMAL);
						}
					} else {	// パスからITEMIDLISTの作成に失敗
						::ShellExecute(NULL, NULL, L"explorer", strPath, NULL, SW_NORMAL);
					}
					::CoTaskMemFree(strPath);
				} else {
				}
#endif
			}
			psi->Release();
		}
    }

private:
    ~CExecuteCommandVerb()
    {
        SafeRelease(&_psia);
        SafeRelease(&_punkSite);
    }

    long _cRef;
    IShellItemArray *_psia;
    IUnknown *_punkSite;
    HWND _hwnd;
    POINT _pt;
    int _nShow;
    DWORD _grfKeyState;
};

// this is called to invoke the verb but this call must not block the caller. to accomidate that
// this function captures the state it needs to invoke the verb and queues a callback via
// the message queue. if your application has a message queue of its own this can be accomilished
// using PostMessage() or setting a timer of zero seconds

IFACEMETHODIMP CExecuteCommandVerb::Execute()
{
    // capture state from the site needed to invoke the verb here
    // note the HWND can be retrieved here but it should not be used for modal UI as
    // all shell verbs should be modeless as well as not block the caller
    IUnknown_GetWindow(_punkSite, &_hwnd);

    // queue the execution of the verb via the message pump
    QueueAppCallback();

    return S_OK;
}

HRESULT CExecuteCommandVerb::Run()
{
    CStaticClassFactory<CExecuteCommandVerb> classFactory(static_cast<IObjectWithSite*>(this));

    HRESULT hr = classFactory.Register(CLSCTX_LOCAL_SERVER, REGCLS_SINGLEUSE);
    if (SUCCEEDED(hr))
    {
        MessageLoop();
    }
    return S_OK;
}

int APIENTRY wWinMain(HINSTANCE, HINSTANCE, PWSTR pszCmdLine, int)
{
	OSVERSIONINFO	osvi = { sizeof(osvi) };
	GetVersionEx(&osvi);
	if ( !(osvi.dwMajorVersion == 6 && osvi.dwMinorVersion == 1) ) {	// Win7以外
		if (StrStrI(pszCmdLine, L"-Register") || StrStrI(pszCmdLine, L"-UnRegister"))
			return 0;
		MessageBox(NULL, L"Windows7以外では使えません", NULL, MB_ICONERROR);
		return 0;
	}

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        DisableComExceptionHandling();
        if (StrStrI(pszCmdLine, L"-Embedding")) {
			assert(FALSE);
            CExecuteCommandVerb *pAppDrop = new (std::nothrow) CExecuteCommandVerb();
            if (pAppDrop) {
                pAppDrop->Run();
                pAppDrop->Release();
            }
        } else if (StrStrI(pszCmdLine, L"-UnRegister")) {
			CRegisterExtension re(__uuidof(CExecuteCommandVerb));
			re.UnRegisterObject();
			UnRegisterExecuteCommandVerb();

		} else if (StrStrI(pszCmdLine, L"-Register")) {
            CRegisterExtension re(__uuidof(CExecuteCommandVerb));
            hr = re.RegisterAppAsLocalServer(c_szVerbDisplayName);
            if (SUCCEEDED(hr))
            {
                //WCHAR const c_szProgID[] = L"Folder";
                //WCHAR const c_szVerbName[] = L"openInTab2";

                //// register this verb on .txt files ProgID
                //hr = re.RegisterExecuteCommandVerb(c_szProgID, c_szVerbName, c_szVerbDisplayName);
                //if (SUCCEEDED(hr))
                //{
                //    hr = re.RegisterVerbAttribute(c_szProgID, c_szVerbName, L"NeverDefault");
                //    if (SUCCEEDED(hr))
                //    {
                //        MessageBox(NULL,
                //            L"Installed ExecuteCommand Verb Sample for .txt files\n\n"
                //            L"right click on a .txt file and choose 'ExecuteCommand Verb Sample' to see this in action",
                //            c_szVerbDisplayName, MB_OK);
                //    }
                //}
            }
        }

        CoUninitialize();
    }

    return 0;
}

































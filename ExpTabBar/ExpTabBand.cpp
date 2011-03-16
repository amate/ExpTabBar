// ExpTabBand.cpp : CExpTabBand �̎���

#include "stdafx.h"

#include "ExpTabBand.h"


// �R���X�g���N�^/�f�X�g���N�^
CExpTabBand::CExpTabBand()
	: m_hWndParent(NULL)
{
}

CExpTabBand::~CExpTabBand()
{
	m_pSink->DispEventUnadvise(m_spWebBrowser2);
}


// IDeskBand
STDMETHODIMP CExpTabBand::GetBandInfo(DWORD /* dwBandID */, DWORD /* dwViewMode */, DESKBANDINFO* pdbi)
{
    if (pdbi)
    {
        //�c�[���o�[�̍ŏ��T�C�Y
        if (pdbi->dwMask & DBIM_MINSIZE)
        {
            pdbi->ptMinSize.x = 100;
            pdbi->ptMinSize.y = 24;
        }
        //�c�[���o�[�̍ő�T�C�Y
        if (pdbi->dwMask & DBIM_MAXSIZE)
        {
            pdbi->ptMaxSize.x = -1;
            pdbi->ptMaxSize.y = -1;
        }
        // �c�[���o�[�ƃr���[�Ƃ̊Ԃ��h���b�O�����Ƃ��ǂꂮ�炢�̊Ԋu���󂯂邩
		// dwModeFlags �� DBIMF_VARIABLEHEIGHT ���ݒ肳��Ă��Ȃ��Ɩ��������
        if (pdbi->dwMask & DBIM_INTEGRAL)
        {
            pdbi->ptIntegral.x = -1;
            pdbi->ptIntegral.y = -1;
        }

        if (pdbi->dwMask & DBIM_ACTUAL)
        {
            pdbi->ptActual.x = -1;
            pdbi->ptActual.y = -1;
        }
        if (pdbi->dwMask & DBIM_TITLE)
        {
            //�c�[���o�[�����ɕ\�������^�C�g��������
			//wcscpy_s(pdbi->wszTitle, _T("ExpTabBar"));
        }
        if( pdbi->dwMask & DBIM_BKCOLOR )
        {
            //���������Ă����΁A�f�t�H���g�̔w�i�F�ɂȂ��Ă����炵��
            pdbi->dwMask &= ~DBIM_BKCOLOR;
        }
        if (pdbi->dwMask & DBIM_MODEFLAGS)
        {
            pdbi->dwModeFlags = DBIMF_NORMAL | DBIMF_USECHEVRON |
                // �o�[��\������ۂɁA�K�����̍s�ɂ��Ă����w��
                DBIMF_BREAK;
        }
    }
    return S_OK;
}


// IOleWindow
STDMETHODIMP CExpTabBand::GetWindow(HWND* phwnd)
{
	HRESULT hr = S_OK;
	if (NULL == phwnd) {
		hr = E_INVALIDARG;
	} else {
		*phwnd = m_wndReflection.GetTabBar().m_hWnd;
	}
	return hr;
}

STDMETHODIMP CExpTabBand::ContextSensitiveHelp(BOOL /* fEnterMode */)
{
    return E_NOTIMPL;
}


// IDockingWindow
STDMETHODIMP CExpTabBand::ShowDW(BOOL fShow)
{
    //�c�[���o�[��ShowWindow�����s����
    //�܂��c�[���o�[������ĂȂ��̂ŁA�������Ȃ�

    return S_OK;
}

STDMETHODIMP CExpTabBand::CloseDW(unsigned long /* dwReserved */)
{
    //CLOSE���̓���B�c�[���o�[���\���ɂ���
    ShowDW(FALSE);

    return S_OK;
}

STDMETHODIMP CExpTabBand::ResizeBorderDW(const RECT* /* prcBorder */, IUnknown* /* punkToolbarSite */, BOOL /* fReserved */)
{
    return E_NOTIMPL;
}

///////////////////////////////////////////////////////////////////////////
//
// IObjectWithSite implementations
//

STDMETHODIMP CExpTabBand::SetSite(IUnknown* punkSite)
{
	HRESULT hr;
	// punkSite��NULL�łȂ��Ȃ�A�V����site���ݒ肳��Ă���
	if(punkSite)
	{
		//Get the parent window.
		CComQIPtr<IOleWindow> pOleWindow(punkSite);
		ATLASSERT(pOleWindow);

		pOleWindow->GetWindow(&m_hWndParent);

		if(m_hWndParent == NULL) {
			ATLASSERT(0);
			return E_FAIL;
		}
		
		// IServiceProvider�̎擾
		CComQIPtr<IServiceProvider> pSP(punkSite);
		ATLASSERT(pSP);

		hr = pSP->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (LPVOID*)&m_spWebBrowser2);
		ATLASSERT(m_spWebBrowser2);

		hr = CComObject<CWebBrowserEvents>::CreateInstance(&m_pSink);
		if (FAILED(hr)) {
			ATLASSERT(FALSE);
			return E_FAIL;
		}

		hr = m_pSink->DispEventAdvise(m_spWebBrowser2);
		if (FAILED(hr)) {
			ATLASSERT(FALSE);
			return E_FAIL;
		}
		m_pSink->SetTabBarWindow(&m_wndReflection.GetTabBar());


		m_wndReflection.GetTabBar().Initialize(punkSite);

		if (RegisterAndCreateWindow() == FALSE) {
			ATLASSERT(FALSE);
			return E_FAIL;
		}
		
		// NOTE: CreateToolWindow call should be last, it relies on m_pIE 
	}

	return S_OK;
}

STDMETHODIMP CExpTabBand::GetSite(REFIID riid, LPVOID *ppvReturn)
{
	return E_FAIL;
}


// Implementation
BOOL	CExpTabBand::RegisterAndCreateWindow()
{
	RECT rect;
	::GetClientRect(m_hWndParent, &rect);
	m_wndReflection.Create(m_hWndParent, rect, NULL, WS_CHILD);
	// The toolbar is the window that the host will be using so it is the window that is important.
	return m_wndReflection.GetTabBar().IsWindow();
}




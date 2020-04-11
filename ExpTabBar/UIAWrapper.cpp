#include "stdafx.h"
#include "UIAWrapper.h"
#include "Misc.h"

UIAWrapper::UIAWrapper()
{
	CString UIADllPath = Misc::GetExeDirectory() + L"UIA.dll";
	m_UIADll = ::LoadLibrary(UIADllPath);
	if (m_UIADll == NULL) {
		return;
	}
	m_funcGetScrollPos = (funcGetScrollPos)::GetProcAddress(m_UIADll, "GetScrollPos");
	ATLASSERT(m_funcGetScrollPos);
	m_funcSetScrollPos = (funcSetScrollPos)::GetProcAddress(m_UIADll, "SetScrollPos");
	ATLASSERT(m_funcSetScrollPos);

	m_funcGetUIItemsViewVerticalScrollPercent 
		= (funcGetUIItemsViewVerticalScrollPercent)::GetProcAddress(m_UIADll, "GetUIItemsViewVerticalScrollPercent");
	ATLASSERT(m_funcGetUIItemsViewVerticalScrollPercent);
	m_funcSetUIItemsViewVerticalScrollPercent 
		= (funcSetUIItemsViewVerticalScrollPercent)::GetProcAddress(m_UIADll, "SetUIItemsViewVerticalScrollPercent");
	ATLASSERT(m_funcSetUIItemsViewVerticalScrollPercent);
}

UIAWrapper::~UIAWrapper()
{
	::FreeLibrary(m_UIADll);
}

int UIAWrapper::GetScrollPos(HWND hwndListView)
{
	int pos = m_funcGetScrollPos(hwndListView);
	return pos;
}

void UIAWrapper::SetScrollPos(HWND hwndListView, int scrollPos)
{
	m_funcSetScrollPos(hwndListView, scrollPos);
}

double UIAWrapper::GetUIItemsViewVerticalScrollPercent(HWND hwndUIItemsView)
{
	double pos = m_funcGetUIItemsViewVerticalScrollPercent(hwndUIItemsView);
	return pos;
}

void UIAWrapper::SetUIItemsViewVerticalScrollPercent(HWND hwndUIItemsView, double virticalScrollPercent)
{
	m_funcSetUIItemsViewVerticalScrollPercent(hwndUIItemsView, virticalScrollPercent);
}

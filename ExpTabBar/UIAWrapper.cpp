#include "stdafx.h"
#include "UIAWrapper.h"
#include "Misc.h"

#define IF_FAILED_THROW(exp) \
	hr = exp;	\
	if (FAILED(hr)) throw hr;

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

	//HRESULT	hr;
	//hr = ::CoCreateInstance(__uuidof(CUIAutomation), NULL, CLSCTX_INPROC_SERVER, __uuidof(IUIAutomation), (void**)&m_spUIAutomation);
	//ATLASSERT(m_spUIAutomation);
}

UIAWrapper::~UIAWrapper()
{
	::FreeLibrary(m_UIADll);
}

int UIAWrapper::GetScrollPos(HWND hwndListView)
{
	int pos = m_funcGetScrollPos(hwndListView);
	return pos;
#if 0
	try {
		HRESULT hr = S_OK;
		CComPtr<IUIAutomationElement>	spElmListView;
		IF_FAILED_THROW(m_spUIAutomation->ElementFromHandle(hwndListView, &spElmListView));
		ATLASSERT(spElmListView);

		CComPtr<IUIAutomationCondition>	spCondScrollbar;
		IF_FAILED_THROW(m_spUIAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, CComVariant(UIA_ScrollBarControlTypeId), &spCondScrollbar));

		CComPtr<IUIAutomationElement>	slElmScrollbar;
		IF_FAILED_THROW(spElmListView->FindFirst(TreeScope_Children, spCondScrollbar, &slElmScrollbar));
		if (!slElmScrollbar) {
			return 0;
		}
		ATLASSERT(slElmScrollbar);

		CComPtr<IUIAutomationRangeValuePattern>	spRangeValue;
		IF_FAILED_THROW(slElmScrollbar->GetCurrentPatternAs(UIA_RangeValuePatternId, IID_IUIAutomationRangeValuePattern, (void**)&spRangeValue));
		ATLASSERT(spRangeValue);

		double value = 0.0;
		IF_FAILED_THROW(spRangeValue->get_CurrentValue(&value));
		int pos = static_cast<int>(value);
		return pos;
	}
	catch (HRESULT hr) {
		ATLASSERT(FALSE);
		return 0;
	}

	//var listViewElm = AutomationElement.FromHandle(hwndListView);
	//var condScrollbarCtrl = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ScrollBar);
	//scrollbarElm = listViewElm.FindFirst(TreeScope.Children, condScrollbarCtrl);
	//var rangeValuePattern = (RangeValuePattern)scrollbarElm.GetCurrentPattern(RangeValuePattern.Pattern);
	//return (int)rangeValuePattern.Current.Value;

	//int pos = m_funcGetScrollPos(hwndListView);
	//return pos;
#endif
}

bool UIAWrapper::SetScrollPos(HWND hwndListView, int scrollPos)
{
	enum { kScrollGap = 24 };
	bool b = m_funcSetScrollPos(hwndListView, scrollPos);
	if (b) {
		//m_funcSetScrollPos(hwndListView, scrollPos + kScrollGap);
	}
	return b;
#if 0
	try {
		HRESULT hr = S_OK;
		CComPtr<IUIAutomationElement>	spElmListView;
		IF_FAILED_THROW(m_spUIAutomation->ElementFromHandle(hwndListView, &spElmListView));
		ATLASSERT(spElmListView);

		CComPtr<IUIAutomationCondition>	spCondScrollbar;
		IF_FAILED_THROW(m_spUIAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, CComVariant(UIA_ScrollBarControlTypeId), &spCondScrollbar));

		CComPtr<IUIAutomationElement>	slElmScrollbar;
		IF_FAILED_THROW(spElmListView->FindFirst(TreeScope_Children, spCondScrollbar, &slElmScrollbar));
		if (!slElmScrollbar) {
			return false;
		}

		CComPtr<IUIAutomationRangeValuePattern>	spRangeValue;
		IF_FAILED_THROW(slElmScrollbar->GetCurrentPatternAs(UIA_RangeValuePatternId, IID_IUIAutomationRangeValuePattern, (void**)&spRangeValue));
		ATLASSERT(spRangeValue);

		IF_FAILED_THROW(spRangeValue->SetValue(scrollPos));
		return true;
	}
	catch (HRESULT hr) {
		ATLASSERT(FALSE);
		return false;
	}

	//m_funcSetScrollPos(hwndListView, scrollPos);
#endif
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

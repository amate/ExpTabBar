#pragma once

class UIAWrapper
{
public:
	UIAWrapper();
	~UIAWrapper();

	int		GetScrollPos(HWND hwndListView);
	void	SetScrollPos(HWND hwndListView, int scrollPos);
	
	double	GetUIItemsViewVerticalScrollPercent(HWND hwndUIItemsView);
	void	SetUIItemsViewVerticalScrollPercent(HWND hwndUIItemsView, double virticalScrollPercent);

private:
	HMODULE	m_UIADll = NULL;

	using funcGetScrollPos = int(*)(HWND);
	funcGetScrollPos	m_funcGetScrollPos = nullptr;
	using funcSetScrollPos = void (*)(HWND, int);
	funcSetScrollPos	m_funcSetScrollPos = nullptr;

	using funcGetUIItemsViewVerticalScrollPercent = double (*)(HWND);
	funcGetUIItemsViewVerticalScrollPercent m_funcGetUIItemsViewVerticalScrollPercent = nullptr;
	using funcSetUIItemsViewVerticalScrollPercent = void (*)(HWND, double);
	funcSetUIItemsViewVerticalScrollPercent m_funcSetUIItemsViewVerticalScrollPercent = nullptr;
};


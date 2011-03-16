// DonutFunc.h
// Donutに関係した他雑な関数

#pragma once

#include "IntShCut.h"

#include "IniFile.h"
#include "Misc.h"


extern TCHAR	g_szIniFileName[MAX_PATH];

// IDとクラス名の配列から子ウィンドウを探す
HWND FindChildWindowIDRecursive(HWND hParent, LPCTSTR* szClass, const int* ID);


// Donut.exeと同じディレクトリに strFile があるものとしてフルパスを返す
CString	_GetFilePath(const CString& strFile);


bool	MtlIsCrossRect(const CRect &rc1, const CRect &rc2);

CString	MtlCompactString(const CString& str, int nMaxTextLength);

// app.exe -> app.ini
void	_IniFileNameInit(LPTSTR lpszIniFileName, DWORD nSize, LPCTSTR lpszExtText = _T(".ini"));

CString _GetSkinDir();


BOOL	_ReplaceImageList(CString strBmpFile, CImageList &imgs, DWORD dfltRes);


int		_Pack(int hi, int low);
BOOL	_QueryColorString(CIniFileRead& pr, COLORREF& col, LPCTSTR lpszKey);


////////////////////////////////////////////////////////////////////////////////////








namespace MTL
{

template <class _Type1, class _Type2>
inline bool _check_flag(_Type1 __flag, _Type2 __flags)
{
	return (__flags & __flag) != 0;
}


/////////////////////////////////////////////////////////////////////////////
// CTrackMouseLeave

template <class T>
class CTrackMouseLeave
{
public:

	// コンストラクタ
	CTrackMouseLeave() : m_bTrackMouseLeave(false) { }

private:
	// Overridables
	void	OnTrackMouseMove(UINT nFlags, CPoint pt) { }
	void	OnTrackMouseLeave() { }


public:
	// メッセージマップ
	BEGIN_MSG_MAP(CTrackMouseLeave<T>)
		MESSAGE_HANDLER(WM_MOUSEMOVE	, OnMouseMove)
		MESSAGE_HANDLER(WM_MOUSELEAVE	, OnMouseLeave)
	END_MSG_MAP()

	// マウスがウィンドウ上を移動したときに呼ばれる
	LRESULT	OnMouseMove(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL &bHandled)
	{
		bHandled = FALSE;
		_StartTrackMouseLeave();
		POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
		T*	pT = static_cast<T *>(this);
		pT->OnTrackMouseMove( (UINT)wParam, pt );
		return 1;
	}

	// マウスがウィンドウから離れたときに呼ばれる
	LRESULT	OnMouseLeave(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
	{
		bHandled = FALSE;
		m_bTrackMouseLeave = false;
		T*	pT = static_cast<T *>(this);
		pT->OnTrackMouseLeave();
		return 0;
	}
private:
	BOOL	_StartTrackMouseLeave()
	{
		if (m_bTrackMouseLeave)
			return FALSE;

		T*	pT = static_cast<T *>(this);
		TRACKMOUSEEVENT tme = { sizeof (tme) };
		tme.dwFlags 	   = TME_LEAVE;
		tme.hwndTrack	   = pT->m_hWnd;
		m_bTrackMouseLeave = true;
		return ::TrackMouseEvent(&tme);
	}
protected:
	bool	m_bTrackMouseLeave;
};
// MtlWin
/////////////////////////////////////////////////////////////////////////////////////



int MtlGetLBTextFixed(HWND hWndCombo, int nIndex, CString &strText);



HDROP MtlCreateDropFile(CSimpleArray<CString> &arrFiles);

// インターネットショートカットファイルを作る
bool MtlCreateInternetShortcutFile(const CString &strFileName, const CString &strUrl);



// ファイル名に使えない文字を置換する
inline void MtlValidateFileName(CString &strName)
{
	strName.Replace( _T("\\"), _T("-") );
	strName.Replace( _T("/" ), _T("-") );
	strName.Replace( _T(":" ), _T("-") );
	strName.Replace( _T("," ), _T("-") );
	strName.Replace( _T(";" ), _T("-") );
	strName.Replace( _T("*" ), _T("-") );
	strName.Replace( _T("?" ), _T("-") );
	strName.Replace( _T("\""), _T("-") );
	strName.Replace( _T("<" ), _T("-") );
	strName.Replace( _T(">" ), _T("-") );
	strName.Replace( _T("|" ), _T("-") );
}


// strPathの拡張子がstrExtかどうかを比較する
inline bool MtlIsExt(const CString &strPath, const CString &strExt)
{
	CString 	strExtSrc = strPath.Right(4);
	return (strExtSrc.CompareNoCase(strExt) == 0);
}


// strの一番後ろとchを比較する
inline bool __MtlIsLastChar(const CString &str, TCHAR ch)
{
	if (str.GetLength() <= 0)
		return false;

	if (str[str.GetLength() - 1] == ch)
		return true;
	else
		return false;
}

// strDirectoryPathの後ろに ch をつける
inline void MtlMakeSureTrailingChar(CString &strDirectoryPath, TCHAR ch)
{
	if ( !__MtlIsLastChar(strDirectoryPath, ch) )
		strDirectoryPath += ch;
}


// strDirectoryPathの後ろに'/'をつける
inline void MtlMakeSureTrailingBackSlash(CString &strDirectoryPath)
{
	MtlMakeSureTrailingChar( strDirectoryPath, _T('\\') );
}


// strTextの幅を計算する
inline int MtlComputeWidthOfText(const CString &strText, HFONT hFont)
{
	if ( strText.IsEmpty() )
		return 0;

	CString str = strText;
	str.Remove(_T('&'));

	CWindowDC dc(NULL);
	CRect	  rcText(0, 0, 0, 0);
	HFONT	  hOldFont = dc.SelectFont(hFont);
	dc.DrawText(str, -1, &rcText, DT_SINGLELINE | DT_LEFT | DT_VCENTER | DT_CALCRECT);
	dc.SelectFont(hOldFont);

	return rcText.Width();
}



template <class _Function>
bool MtlForEachFile(const CString &strDirectoryPath, _Function __f)
{
	CString 		strPathFind = strDirectoryPath;

	MtlMakeSureTrailingBackSlash(strPathFind);
	CString 		strPath 	= strPathFind;
	strPathFind += _T("*.*");

	WIN32_FIND_DATA wfd;
	HANDLE			h	= ::FindFirstFile(strPathFind, &wfd);

	if (h == INVALID_HANDLE_VALUE)
		return false;

	// Now scan the directory
	do {
		// it is a file
		if ( ( wfd.dwFileAttributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_HIDDEN | FILE_ATTRIBUTE_SYSTEM) ) == 0 ) {
			__f(strPath + wfd.cFileName);	// 関数オブジェクト__fにファイルパスを渡す
		}
	} while ( ::FindNextFile(h, &wfd) );

	::FindClose(h);

	return true;
}

















}	// namespace MTL

using namespace MTL;


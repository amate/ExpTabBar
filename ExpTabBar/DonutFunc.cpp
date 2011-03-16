// DonutFunc.cpp

#include "stdafx.h"

#include "DonutFunc.h"

#include "Misc.h"



// IDとクラス名の配列から子ウィンドウを探す
/*	// 使用例
    LPSTR kls[] = {"DUIViewWndClassName","DirectUIHWND","CtrlNotifySink","SHELLDLL_DefView",NULL};
    const int ids[] = {0,0,0,0x461,-1};
    hWnd = FindChildWindowIDRecursive(hDlg, kls , ids);
*/
HWND FindChildWindowIDRecursive(HWND hParent, LPCTSTR* szClass, const int* ID)
{
	HWND hWnd = NULL;
	while (1) {
		hWnd = ::FindWindowEx( hParent, hWnd, *szClass, NULL);

		if (hWnd == NULL) {
			return NULL;		// 子ウィンドウが見つからなかった
		}

		int tID = ::GetWindowLong( hWnd, GWL_ID);
		if( tID == *ID){
			ATLTRACE(" IDがマッチした\n");
			if (szClass + 1 == NULL || *(ID + 1) == -1) {
				ATLTRACE("見つかりました\n");
				return hWnd;	// 最後まで行ったので終わり
			} else {
				HWND hWndChild = FindChildWindowIDRecursive(hWnd, szClass + 1, ID + 1);
				if(hWndChild != NULL){
					return hWndChild;
				}
			}
		}
	ATLTRACE(" IDがマッチしなかった\n");
	}
	return NULL;
}


CString	_GetFilePath(const CString& strFile)
{
	return Misc::GetExeDirectory() + strFile;
}



bool	MtlIsCrossRect(const CRect &rc1, const CRect &rc2)
{
	CRect	rcCross = rc1 & rc2;
	
	return rcCross.IsRectEmpty() == 0;
}

CString	MtlCompactString(const CString& str, int nMaxTextLength)
{
	ATLASSERT(nMaxTextLength > 0);

	int	l = Misc::eseHankakuStrLen(str);

	if (l <= nMaxTextLength)
		return str;

	LPCTSTR	szEllipsis	= _T("...");
	const int cchEndEllipsis = 3;

	if (nMaxTextLength <= cchEndEllipsis) {
		return CString(_T('.'), nMaxTextLength);
	}

	int	nIndex = nMaxTextLength - cchEndEllipsis;
	ATLASSERT(nIndex >= 0);

	return Misc::eseHankakuStrLeft(str, nIndex) + szEllipsis;
}


// app.exe -> app.ini
void _IniFileNameInit(LPTSTR lpszIniFileName, DWORD nSize, LPCTSTR lpszExtText)
{
	LPTSTR	lpszExt = NULL;
	LPTSTR	lpsz;

	::GetModuleFileName(g_hInst, lpszIniFileName, nSize);

	for ( lpsz = lpszIniFileName ; *lpsz != NULL ; lpsz = ::CharNext(lpsz) ) {
		if ( *lpsz == _T('.') )		// 最後の '.' を探す
			lpszExt = lpsz;
	}

	ATLASSERT(::lstrlen(lpszExtText) == 4);
	::lstrcpy(lpszExt, lpszExtText);
}

CString _GetSkinDir()
{
	CString strSkinPath;		// 最終的なスキンのパスが入る
	CString	strPath		= Misc::GetExeDirectory() + _T("Skin\\");
	CString	strDefPath	= strPath + _T("default");

	CIniFileRead	pr( g_szIniFileName, _T("Skin") );
	strSkinPath = pr.GetStringUW( _T("SkinFolder") );
	pr.Close();

	if ( strSkinPath.IsEmpty() )
		strSkinPath = strDefPath;
	else
		strSkinPath = strPath + strSkinPath;

	// フォルダの存在チェック
	if (::GetFileAttributes(strSkinPath) == 0xFFFFFFFF) {
		if (::GetFileAttributes(strDefPath) == 0xFFFFFFFF) {
			strSkinPath = strPath;
		} else {
			strSkinPath = strDefPath;
		}
	}

	// '\'が最後についていなければつける
	if ( !strSkinPath.IsEmpty() && strSkinPath.Right(1) != _T("\\") )
		strSkinPath += _T("\\");

	return strSkinPath;
}

BOOL _ReplaceImageList(CString strBmpFile, CImageList &imgs, DWORD dfltRes)
{
	CBitmap 	bmp;
	bmp.Attach( AtlLoadBitmapImage(strBmpFile.GetBuffer(0), LR_LOADFROMFILE) );
	if (bmp.m_hBitmap == 0 && dfltRes) {
		// ファイルがなかったとき、デフォルトのリソースを読み込んでみる.
		bmp.LoadBitmap(dfltRes);
	}
	if (bmp.m_hBitmap) {
		int nCount = imgs.GetImageCount();
		for (int i = 0; i < nCount; i++) {
			if ( !imgs.Remove(0) )
				return FALSE;
		}
		imgs.Add( bmp, RGB(255, 0, 255) );
		return TRUE;
	}
	return FALSE;			// 失敗
}


static int		_Pack(int hi, int low)
{
	if ( !( ( ('0' <= low && low <= '9') || ('A' <= low && low <= 'F') || ('a' <= low && low <= 'f') )
		  && ( ('0' <= hi && hi  <= '9') || ('A' <= hi	&& hi  <= 'F') || ('a' <= hi  && hi  <= 'f') ) ) )
		return 0;	// 数値ではない

	int nlow = ('0' <= low && low <= '9') ? low - '0'
			 : ('A' <= low && low <= 'F') ? low - 'A' + 0xA
			 :								low - 'a' + 0xA ;
	int nhi  = ('0' <= hi  && hi  <= '9') ? hi	- '0'
			 : ('A' <= hi  && hi  <= 'F') ? hi	- 'A' + 0xA
			 :								hi	- 'a' + 0xA ;

	return (nhi << 4) + nlow;
}

BOOL	_QueryColorString(CIniFileRead& pr, COLORREF& col, LPCTSTR lpszKey)
{
	CString strCol = pr.GetString(lpszKey, NULL, 20);
	// 元々 pr.QueryValue(文字列)でのエラーは、長さが0のときのことなので、文字列が空かどうかのチェックだけで充分
	if ( strCol.IsEmpty() )
		return FALSE;

	strCol.TrimLeft('#');

	col = RGB( _Pack(strCol[0], strCol[1]) ,
			   _Pack(strCol[2], strCol[3]) ,
			   _Pack(strCol[4], strCol[5])
			 );
	return TRUE;
}



namespace MTL {


int MtlGetLBTextFixed(HWND hWndCombo, int nIndex, CString &strText)
{
	CComboBox	combo(hWndCombo);
	ATLASSERT( combo.IsWindow() );
	int 	nRet = combo.GetLBTextLen(nIndex) * 2;
	if (nRet > 0) {
		nRet = combo.GetLBText( nIndex, strText.GetBufferSetLength( nRet+1 ) );
		strText.ReleaseBuffer();
	} else {
		strText.Empty();
	}
	return nRet;
}



HDROP MtlCreateDropFile(CSimpleArray<CString> &arrFiles)
{
	if (arrFiles.GetSize() == 0)
		return NULL;

	//filename\0...\0filename\0...\0filename\0\0
	int 	nLen  = 0;
	int 	i;
	for (i = 0; i < arrFiles.GetSize(); ++i) {
		nLen += arrFiles[i].GetLength();
		nLen += 1;							// for '\0' separator
	}

	nLen	+= 1;				// for the last '\0'

	HDROP	hDrop = (HDROP) ::GlobalAlloc( GHND, sizeof (DROPFILES) + nLen * sizeof (TCHAR) );
	if (hDrop == NULL)
		return NULL;

	LPDROPFILES lpDropFiles;
	lpDropFiles 		= (LPDROPFILES) ::GlobalLock(hDrop);
	lpDropFiles->pFiles = sizeof (DROPFILES);
	lpDropFiles->pt.x	= 0;
	lpDropFiles->pt.y	= 0;
	lpDropFiles->fNC	= FALSE;
	lpDropFiles->fWide	= TRUE;

	TCHAR * 	psz   = (TCHAR *) (lpDropFiles + 1);

	for (i = 0; i < arrFiles.GetSize(); ++i) {
		::lstrcpy(psz, arrFiles[i]);
		psz += arrFiles[i].GetLength() + 1; 	// skip a '\0' separator
	}
	::GlobalUnlock(hDrop);

	return hDrop;
}


bool MtlCreateInternetShortcutFile(const CString &strFileName, const CString &strUrl)
{
	CComPtr<IUniformResourceLocator> spUrl;
	HRESULT	hr = ::CoCreateInstance(CLSID_InternetShortcut, NULL, CLSCTX_ALL, IID_IUniformResourceLocator, (void **) &spUrl);
	if ( FAILED(hr) )
		return false;

	CComPtr<IPersistFile>			 spPf;
	hr = spUrl->QueryInterface(IID_IPersistFile, (void **) &spPf);
	if ( FAILED(hr) )
		return false;

	hr = spUrl->SetURL(strUrl, 0);
	if ( FAILED(hr) )
		return false;
	
	hr = spPf->Save((LPCTSTR)strFileName, TRUE);
	if ( FAILED(hr) )
		return false;

	return true;
}













}	// namespace MTL
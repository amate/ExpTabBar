// IniFile.cpp

#include "stdafx.h"

#include "IniFile.h"

#include <windows.h>

#include <atlbase.h>
#include <atlapp.h>
#include <atlmisc.h>

#include <atlsync.h>

#include "Misc.h"

////////////////////////////////////////////////////////////////////////////////////////
// CIniFile

void	CIniFile::Open(const CString& strFileName, const CString& strSectionName)
{
	ATLASSERT( !strFileName.IsEmpty() );
	ATLASSERT( !IsOpen() );					// もうすでにOpenされていれば止める

	m_strFileName		= strFileName;
	m_strSectionName	= strSectionName;
}

BOOL	CIniFile::IsOpen()
{
	return !m_strFileName.IsEmpty();
}

void	CIniFile::Close()
{
	m_strFileName.Empty();
	m_strSectionName.Empty();
}

void	CIniFile::ChangeSectionName(LPCTSTR sectionName)
{
	ATLASSERT( !m_strFileName.IsEmpty() );

	m_strSectionName = sectionName;
}


////////////////////////////////////////////////////////////////////////////////////////
// CIniFileRead

// コンストラクタ
CIniFileRead::CIniFileRead(const CString& strFileName, const CString& strSectionName)
{
	Open(strFileName, strSectionName);
}

// ValueNameの持つ値を返す、エラーの時、dwValueの値は元のまま
LONG	CIniFileRead::QueryValue(DWORD& dwValue, LPCTSTR lpszValueName)
{
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	UINT nValue = ::GetPrivateProfileInt(m_strSectionName, lpszValueName, cnt_nDefault, m_strFileName);
	if (nValue == cnt_nDefault)
		return ERROR_CANTREAD;

	dwValue = nValue;
	return ERROR_SUCCESS;
}

LONG	CIniFileRead::QueryValue(int& nValue, LPCTSTR lpszValueName)
{
	return QueryValue(*(DWORD*)&nValue, lpszValueName);
}

LONG	CIniFileRead::QueryString(LPTSTR szValue, LPCTSTR lpszValueName, DWORD* pdwCount)
{
	ATLASSERT(pdwCount != NULL);
	ATLASSERT(szValue != NULL );
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	DWORD dw = ::GetPrivateProfileString(m_strSectionName, lpszValueName, _T(""), szValue, *pdwCount, m_strFileName);
	if (dw == 0)
		return ERROR_CANTREAD;

	return ERROR_SUCCESS;
}

// エラーを返す意味がない場合も多いので、エラー時用デフォルト指定ありで、値そのものを返すバージョンを用意
DWORD	CIniFileRead::GetValue(LPCTSTR lpszValueName, DWORD defaultValue)
{
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	return ::GetPrivateProfileInt(m_strSectionName, lpszValueName, defaultValue, m_strFileName);
}

const CString CIniFileRead::GetString(LPCTSTR lpszValueName, LPCTSTR pszDefault, DWORD dwBufSize)
{
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	// バッファサイズが0なら固定で512byte分取得する
	if (dwBufSize == 0)
		dwBufSize = 512;

	TCHAR*	buf = new TCHAR[dwBufSize];
	DWORD	dw = ::GetPrivateProfileString(m_strSectionName, lpszValueName, _T(""), buf, dwBufSize, m_strFileName);
	if (dw == 0 && pszDefault) {			// 文字列が取得できなかったらデフォルトの値を返す
		delete [] buf;
		return CString(pszDefault);
	}
	
	buf[dw] = '\0';
	CString str(buf);
	delete [] buf;
	return str;
}

// 文字列取得。もし、文字列が半角英数記号のみで %?? が混ざっている場合、UTF8文字列をエンコードしたものとして、デコードして返す
const CString CIniFileRead::GetStringUW(LPCTSTR lpszValueName, LPCTSTR pszDefault, DWORD dwBufSize)
{
	return Misc::urlstr_decodeWhenASC(GetString(lpszValueName, pszDefault, dwBufSize));
}


//////////////////////////////////////////////////////////////////////////////////////////
// CIniFileWrite

// コンストラクタ
CIniFileWrite::CIniFileWrite(const CString& strFileName, const CString& strSectionName)
{
	Open(strFileName, strSectionName);
}

LONG	CIniFileWrite::SetValue(DWORD dwValue, LPCTSTR lpszValueName)
{
	ATLASSERT( lpszValueName != NULL );
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	CString strValue;
	strValue.Format(_T("%d"), dwValue);

	if ( ::WritePrivateProfileString(m_strSectionName, lpszValueName, strValue, m_strFileName) )
		return ERROR_SUCCESS;
	else
		return ERROR_CANTWRITE;
}

LONG	CIniFileWrite::SetString(LPCTSTR lpszValue, LPCTSTR lpszValueName)
{
	ATLASSERT( lpszValue != NULL );
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	ATLASSERT( m_strFileName.GetLength() < 4095 );

	if ( ::WritePrivateProfileString(m_strSectionName, lpszValueName, lpszValue, m_strFileName) )
		return ERROR_SUCCESS;
	else
		return ERROR_CANTWRITE;
}

LONG	CIniFileWrite::SetStringUW(LPCTSTR lpszValue, LPCTSTR lpszValueName)
{
	MessageBox( NULL, _T("まだ実装していません"), NULL, NULL );
	return ERROR_CANTWRITE;
}

LONG	CIniFileWrite::DeleteValue(LPCTSTR lpszValueName)
{
	ATLASSERT( lpszValueName != NULL );
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	if ( ::WritePrivateProfileString(m_strSectionName, lpszValueName, NULL, m_strFileName) )
		return ERROR_SUCCESS;
	else
		return ERROR_CANTWRITE;
}

bool	CIniFileWrite::DeleteSection()
{
	ATLASSERT( IsOpen() );
	ATLASSERT( !m_strSectionName.IsEmpty() );

	return ::WritePrivateProfileString(m_strSectionName, NULL, NULL, m_strFileName) != 0;
}


/////////////////////////////////////////////////////////////////////////////////////////
// CIniFileIO

// コンストラクタ
CIniFileIO::CIniFileIO(const CString& strFileName, const CString& strSectionName)
{
	CIniFile::Open(strFileName, strSectionName);
}

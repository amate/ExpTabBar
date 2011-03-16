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
	ATLASSERT( !IsOpen() );					// �������ł�Open����Ă���Ύ~�߂�

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

// �R���X�g���N�^
CIniFileRead::CIniFileRead(const CString& strFileName, const CString& strSectionName)
{
	Open(strFileName, strSectionName);
}

// ValueName�̎��l��Ԃ��A�G���[�̎��AdwValue�̒l�͌��̂܂�
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

// �G���[��Ԃ��Ӗ����Ȃ��ꍇ�������̂ŁA�G���[���p�f�t�H���g�w�肠��ŁA�l���̂��̂�Ԃ��o�[�W������p��
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

	// �o�b�t�@�T�C�Y��0�Ȃ�Œ��512byte���擾����
	if (dwBufSize == 0)
		dwBufSize = 512;

	TCHAR*	buf = new TCHAR[dwBufSize];
	DWORD	dw = ::GetPrivateProfileString(m_strSectionName, lpszValueName, _T(""), buf, dwBufSize, m_strFileName);
	if (dw == 0 && pszDefault) {			// �����񂪎擾�ł��Ȃ�������f�t�H���g�̒l��Ԃ�
		delete [] buf;
		return CString(pszDefault);
	}
	
	buf[dw] = '\0';
	CString str(buf);
	delete [] buf;
	return str;
}

// ������擾�B�����A�����񂪔��p�p���L���݂̂� %?? ���������Ă���ꍇ�AUTF8��������G���R�[�h�������̂Ƃ��āA�f�R�[�h���ĕԂ�
const CString CIniFileRead::GetStringUW(LPCTSTR lpszValueName, LPCTSTR pszDefault, DWORD dwBufSize)
{
	return Misc::urlstr_decodeWhenASC(GetString(lpszValueName, pszDefault, dwBufSize));
}


//////////////////////////////////////////////////////////////////////////////////////////
// CIniFileWrite

// �R���X�g���N�^
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
	MessageBox( NULL, _T("�܂��������Ă��܂���"), NULL, NULL );
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

// �R���X�g���N�^
CIniFileIO::CIniFileIO(const CString& strFileName, const CString& strSectionName)
{
	CIniFile::Open(strFileName, strSectionName);
}

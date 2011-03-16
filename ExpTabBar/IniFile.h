// IniFile.h
#pragma once

// �����ӁA�t�@�C�����̓t���p�X�ŗ^���Ȃ��ƃt�@�C����Windows�f�B���N�g���ӂ�ɐ��������

///////////////////////////////////////////////////////////////////////////////////
// ���N���X
class CIniFile
{
protected:
	enum { cnt_nDefault = 0xABCD0123 }; // �f�t�H���g�l(2882339107)

	CString	m_strFileName;
	CString	m_strSectionName;
public:
	void	Open(const CString& strFileName, const CString& strSectionName);
	BOOL	IsOpen();			// �J���Ă����TRUE���A��

	virtual void Close();

	void	ChangeSectionName(LPCTSTR sectionName);		// �Z�N�V�������̕ύX
};

////////////////////////////////////////////////////////////////////////////////////
// .ini�̓ǂݍ��݂��s�����߂̃N���X
class CIniFileRead : public CIniFile
{
public:
	// �R���X�g���N�^
	CIniFileRead(){};
	CIniFileRead(const CString& strFileName, const CString& strSectionName);

	LONG	QueryValue(DWORD& dwValue, LPCTSTR lpszValueName);
	LONG	QueryValue(int& nValue	 , LPCTSTR lpszValueName);
	LONG	QueryString(LPTSTR szValue, LPCTSTR lpszValueName, DWORD* pdwCount);

	DWORD	GetValue(LPCTSTR lpszValueName, DWORD defaultValue = 0);
	const CString GetString(LPCTSTR lpszValueName, LPCTSTR pszDefault = 0, DWORD dwBufSize = 0);
	const CString GetStringUW(LPCTSTR lpszValueName, LPCTSTR pszDefault = 0, DWORD dwBufSize = 0);
};

////////////////////////////////////////////////////////////////////////////////////
// .ini�̏������݂��s�����߂̃N���X

class CIniFileWrite : public CIniFile
{
public:
	// �R���X�g���N�^
	CIniFileWrite(){};
	CIniFileWrite(const CString& strFileName, const CString& strSectionName = _T(""));

	LONG	SetValue(DWORD dwValue, LPCTSTR lpszValueName);
	LONG	SetString(LPCTSTR lpszValue, LPCTSTR lpszValueName);
	LONG	SetStringUW(LPCTSTR lpszValue, LPCTSTR lpszValueName);

	LONG	DeleteValue(LPCTSTR lpszValueName);
	bool	DeleteSection();						// �Z�N�V�����������Ă��珑���ꍇ�p

};

/////////////////////////////////////////////////////////////////////////////////////
// .ini�̓ǂݏ������s�����߂̃N���X
class CIniFileIO : public CIniFileRead, public CIniFileWrite
{
public:
	// �R���X�g���N�^
	CIniFileIO(const CString& strFileName, const CString& strSectionName = _T(""));
};
// IniFile.h
#pragma once

// ※注意、ファイル名はフルパスで与えないとファイルはWindowsディレクトリ辺りに生成される

///////////////////////////////////////////////////////////////////////////////////
// 基底クラス
class CIniFile
{
protected:
	enum { cnt_nDefault = 0xABCD0123 }; // デフォルト値(2882339107)

	CString	m_strFileName;
	CString	m_strSectionName;
public:
	void	Open(const CString& strFileName, const CString& strSectionName);
	BOOL	IsOpen();			// 開いていればTRUEが帰る

	virtual void Close();

	void	ChangeSectionName(LPCTSTR sectionName);		// セクション名の変更
};

////////////////////////////////////////////////////////////////////////////////////
// .iniの読み込みを行うためのクラス
class CIniFileRead : public CIniFile
{
public:
	// コンストラクタ
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
// .iniの書き込みを行うためのクラス

class CIniFileWrite : public CIniFile
{
public:
	// コンストラクタ
	CIniFileWrite(){};
	CIniFileWrite(const CString& strFileName, const CString& strSectionName = _T(""));

	LONG	SetValue(DWORD dwValue, LPCTSTR lpszValueName);
	LONG	SetString(LPCTSTR lpszValue, LPCTSTR lpszValueName);
	LONG	SetStringUW(LPCTSTR lpszValue, LPCTSTR lpszValueName);

	LONG	DeleteValue(LPCTSTR lpszValueName);
	bool	DeleteSection();						// セクション消去してから書く場合用

};

/////////////////////////////////////////////////////////////////////////////////////
// .iniの読み書きを行うためのクラス
class CIniFileIO : public CIniFileRead, public CIniFileWrite
{
public:
	// コンストラクタ
	CIniFileIO(const CString& strFileName, const CString& strSectionName = _T(""));
};
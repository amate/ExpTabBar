// Misc.cpp

#include "stdafx.h"

#include "Misc.h"


namespace Misc {



HBITMAP	CreateBitmapFromHICON(HICON hIcon)
{
	UINT         uWidth, uHeight;
	HBITMAP      hbmp, hbmpPrev;

	uWidth	= GetSystemMetrics(SM_CXSMICON);
	uHeight = GetSystemMetrics(SM_CYSMICON);

	/* CreateBitmapARGB */
	LPVOID           lpBits;
	BITMAPINFO       bmi = { 0 };
	BITMAPINFOHEADER bmiHeader = { 0 };

	bmiHeader.biSize      = sizeof(BITMAPINFOHEADER);
	bmiHeader.biWidth     = uWidth;
	bmiHeader.biHeight    = uHeight;
	bmiHeader.biPlanes    = 1;
	bmiHeader.biBitCount  = 32;

	bmi.bmiHeader = bmiHeader;
	
	hbmp = CreateDIBSection(NULL, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS, &lpBits, NULL, 0);

	HDC hdcMem = CreateCompatibleDC(NULL);
	hbmpPrev = (HBITMAP)SelectObject(hdcMem, hbmp);
	DrawIconEx(hdcMem, 0, 0, hIcon, uWidth, uHeight, 0, NULL, DI_NORMAL);
	SelectObject(hdcMem, hbmpPrev);
	DeleteDC(hdcMem);

	return hbmp;
#if 0
	try {
		UINT        uWidth, uHeight;
		UINT        uLineSize, uBufferSize;
		LPBYTE      lpBits;
		HBITMAP     hbmp;
		HRESULT		hr;
			
		CComPtr<IWICImagingFactory> spWICImagingFactory;
		hr = spWICImagingFactory.CoCreateInstance(CLSID_WICImagingFactory);
		if (FAILED(hr))
			AtlThrow(hr);

		CComPtr<IWICBitmap>	spWICBitmap;
		hr = spWICImagingFactory->CreateBitmapFromHICON(hIcon, &spWICBitmap);
		if (FAILED(hr))
			AtlThrow(hr);

		hr = spWICBitmap->GetSize(&uWidth, &uHeight);
		if (FAILED(hr))
			AtlThrow(hr);

		/* CreateBitmapARGB */
		BITMAPINFO       bmi = { 0 };
		BITMAPINFOHEADER bmiHeader = { 0 };

		bmiHeader.biSize      = sizeof(BITMAPINFOHEADER);
		bmiHeader.biWidth     = uWidth;
		bmiHeader.biHeight    = -uHeight;
		bmiHeader.biPlanes    = 1;
		bmiHeader.biBitCount  = 32;

		bmi.bmiHeader = bmiHeader;
		
		hbmp = CreateDIBSection(NULL, (LPBITMAPINFO)&bmi, DIB_RGB_COLORS, (void **)&lpBits, NULL, 0);


		uLineSize = uWidth * 4;
		uBufferSize = uLineSize * uHeight;

		hr = spWICBitmap->CopyPixels(NULL, uLineSize, uBufferSize, lpBits);
		if (FAILED(hr))
			AtlThrow(hr);

		return hbmp;
	}
	catch (const CAtlException &e) {
		ATLASSERT(FALSE);
		return NULL;
	}
#endif
}






/** �ȈՂȁA���p�������J�E���g�B
 *	UNICODE�̏ꍇ�A�P���ȕ����R�[�h����ŁA�\�������̔��p�������Z�̒l�ŕ�����T�C�Y��Ԃ�.
 *  ���p�S�p�̔��ʂ͎蔲���K���B���F�t�H���g����Ȃ̂ŁA���\�R�ɂȂ邪�ASJIS��ł�����ۂ��Ȃ�悤��...���炢.
 *  ��؂̊T�Z�����߂�ꍇ�p��.
 */
int	eseHankakuStrLen(const TCHAR* s)
{
  #ifdef UNICODE
	unsigned l = 0;
	while (*s) {
		unsigned c = *s++;
		if (c < 0x200 || (c >= 0xFF61 && c <= 0xFF9F))
			++l;
		else
			l += 2;
	}
	return l;
  #else	//+++ SJIS�Ȃ炻�̂܂ܕ������ł�����
	return strlen(s);
  #endif
}

/** ���񔼊p�w��Ŏw�肳�ꂽ�������܂ł̕������Ԃ�.
 */
const CString eseHankakuStrLeft(const CString& str, unsigned len)
{
  #ifdef UNICODE
	LPCTSTR		t = LPCTSTR(str);
	LPCTSTR		s = t;
	unsigned l = 0;
	unsigned b = 0;
	while (*s) {
		unsigned c = *s;
		if (c < 0x200 || (c >= 0xFF61 && c <= 0xFF9F))
			++l;
		else
			l += 2;
		if (l > len) {
			l = b;
			break;
		}
		b = l;
		++s;
	}
	l = int(s - t);
	return str.Left(l);
  #else	//+++ SJIS�Ȃ炻�̂܂ܕ������ł�����
	return str.Left(len);
  #endif
}





// Donut.exe�̃t���p�X����Ԃ�
const CString GetExeFileName()
{
	TCHAR	buf[MAX_PATH];
	buf[0] = 0;
	::GetModuleFileName(g_hInst, buf, MAX_PATH);
	return CString(buf);
}

// exe�̂���t�H���_��Ԃ�(�Ō�� \ ����)
const CString GetExeDirectory()
{
	CString str = GetExeFileName();
	int		n	= str.ReverseFind( _T('\\') );
	return str.Left(n+1);
}

// exe�̂���t�H���_��Ԃ�(�Ō�� \ ���t���Ȃ�)
const CString GetExeDirName()
{
	CString str = GetExeFileName();
	int		n	= str.ReverseFind( _T('\\') );
	return str.Left(n);
}


/// path ����g���q�𓾂�('.'�͊܂܂ꂸ)
CString GetPathExtention(const CString& path)
{
	LPWSTR strtemp = ::PathFindExtension(path);
	if (strtemp) {
		++strtemp;	// '.' ���΂�
		return strtemp;
	}
	return CString();
}


// UTF8�������wcs(UTF16LE)������ɕϊ�
const std::vector<wchar_t> utf8_to_wcs(const char* pUtf8)
{
	std::vector<wchar_t> vecWcs;
	unsigned chrs = ::MultiByteToWideChar(CP_UTF8, 0, pUtf8, -1, NULL, 0);
	if (chrs) {
		vecWcs.resize( chrs );
		chrs = ::MultiByteToWideChar(CP_UTF8, 0, pUtf8, -1, &vecWcs[0], chrs);
		if ( chrs == 0 )
			vecWcs.clear();
	}
	return vecWcs;
}

const std::vector<TCHAR> utf8_to_tcs(const char* pUtf8)
{
	return utf8_to_wcs(pUtf8);
}

const CString utf8_to_CString(const std::vector<char>& utf8)
{
	return CString(&utf8_to_tcs(&utf8[0])[0]);
}

// %20 �����g��ꂽURL��ʏ�̕�����ɕϊ� ���ɑS�p�������������Ă��Ȃ����ƁI
const std::vector<char> urlstr_decode(const TCHAR* url)
{
	std::vector<char>	buf;

	if (url == NULL)
		return buf;

	const TCHAR* s = url;
	unsigned	 l = 0;
	int			 c;

	while((c = *s++) != 0) {
		++l;
		if ( c == '%' ) {
			if (*s) {
				++s;
				if (*s)
					++s;
			}
		}
	}

	++l;
	buf.resize(l);
	s = url;
	unsigned char* d = (unsigned char*) &buf[0];
	while ((c = *s++) != 0) {
		if (c == '%') {
			c = *s++;
			if (c) {
				if      (c >= '0' && c <= '9') c = c - '0';
				else if (c >= 'A' && c <= 'F') c = 10 + c - 'A';
				else if (c >= 'a' && c <= 'f') c = 10 + c - 'a';
				else                           c = -1;
				int k = *s++;
				if (k && c >= 0) {
					if      (k >= '0' && k <= '9') k = k - '0';
					else if (k >= 'A' && k <= 'F') k = 10 + k - 'A';
					else if (k >= 'a' && k <= 'f') k = 10 + k - 'a';
					else                           k = -1;
					if (k >= 0) {
						c    = (c << 4) | k;
						*d++ = c;
					}
				}
			}
		} else {
			*d++ = char(c);
		}
	}
	*d = '\0';
	return buf;
}


// �����񂪔��p�p���L���݂̂ł���%���������Ă���ꍇ�AUTF8���G���R�[�h�������̂Ƃ��ăf�R�[�h����
const CString urlstr_decodeWhenASC(const CString& str)
{
	// '%' ������Β����Ώ�
	if ((int)str.Find(_T("%")) >= 0) {
		const TCHAR *s = LPCTSTR(str); 
		bool flag = true;

		while (*s) {
			int c = *s++;
			if ((unsigned)c >= 0x80) {	// 0x80�ȏ゠���SJIS�┼�p�J�i���낤�ŁA�ϊ����Ȃ�
				flag = false;
				break;
			}
			if (c == _T('%')) {			// % �������� %?? �̌`�łȂ��Ȃ�A�ϊ����Ȃ�
				if ( !isxdigit(*s) || !isxdigit(s[1]) ) {
					flag = false;
					break;
				}
			}
		}

		if (flag) {						// �p���L���݂̂ł��� %?? ���������Ă��鎞
			return utf8_to_CString( urlstr_decode(str) );
		}
	}
	return str;
}

/// �����Ŏw�肵���E�B���h�E���őO�ʕ\������
void SetForegroundWindow(HWND hWnd)
{
	// �őO�ʃv���Z�X�̃X���b�hID���擾���� 
	int foregroundID = ::GetWindowThreadProcessId( ::GetForegroundWindow(), NULL); 
	// �őO�ʃA�v���P�[�V�����̓��͏����@�\�ɐڑ����� 
	AttachThreadInput( ::GetCurrentThreadId(), foregroundID, TRUE); 
	// �őO�ʃE�B���h�E��ύX���� 
	//::SetForegroundWindow(hWnd);
	::BringWindowToTop(hWnd);

	// �ڑ�����������
	AttachThreadInput( ::GetCurrentThreadId(), foregroundID, FALSE);
#if 0	//\\+
	::SetWindowPos(hWnd, HWND_TOPMOST  , 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSENDCHANGING | SWP_NOSIZE);
	::SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSENDCHANGING | SWP_NOSIZE);
	::SetForegroundWindow(hWnd);						//�E�C���h�E�Ƀt�H�[�J�X���ڂ�
#endif
}



}	// namespace Misc
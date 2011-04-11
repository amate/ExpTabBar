// Misc.h
#pragma once

#include <vector>

#if 0
#include <wincodec.h>
#include <wincodecsdk.h>
#pragma comment(lib, "WindowsCodecs.lib")
#endif

extern HINSTANCE g_hInst;	// ���݂̃p�X���擾���邽�߂ɕK�v

namespace Misc {


HBITMAP	CreateBitmapFromHICON(HICON hIcon);




// �ȈՂȁA���p�������Z�ł̕��������ǂ��J�E���g�B
int	eseHankakuStrLen(const TCHAR* s);

// ���񔼊p�w��Ŏw�肳�ꂽ�������܂ł̕������Ԃ�.
const CString eseHankakuStrLeft(const CString& str, unsigned len);


// Donut.exe�̃t���p�X����Ԃ�
const CString GetExeFileName();

// exe�̂���t�H���_��Ԃ�(�Ō�� \ ����)
const CString GetExeDirectory();

// exe�̂���t�H���_��Ԃ�(�Ō�� \ ���t���Ȃ�)
const CString GetExeDirName();


// UTF8�������wcs(UTF16LE)������ɕϊ�
const std::vector<wchar_t> utf8_to_wcs(const char* pUtf8);

const std::vector<TCHAR> utf8_to_tcs(const char* pUtf8);

const CString utf8_to_CString(const std::vector<char>& utf8);

// %20 �����g��ꂽURL��ʏ�̕�����ɕϊ� ���ɑS�p�������������Ă��Ȃ����ƁI
const std::vector<char> urlstr_decode(const TCHAR* url);


// �����񂪔��p�p���L���݂̂ł���%���������Ă���ꍇ�AUTF8���G���R�[�h�������̂Ƃ��ăf�R�[�h����
const CString urlstr_decodeWhenASC(const CString& str);


/// �����Ŏw�肵���E�B���h�E���őO�ʕ\������
void	SetForegroundWindow(HWND hWnd);


}	// namespace Misc

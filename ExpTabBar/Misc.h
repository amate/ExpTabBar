// Misc.h
#pragma once

#include <vector>

#if 0
#include <wincodec.h>
#include <wincodecsdk.h>
#pragma comment(lib, "WindowsCodecs.lib")
#endif

extern HINSTANCE g_hInst;	// 現在のパスを取得するために必要

namespace Misc {


HBITMAP	CreateBitmapFromHICON(HICON hIcon);




// 簡易な、半角文字換算での文字数もどきカウント。
int	eseHankakuStrLen(const TCHAR* s);

// 似非半角指定で指定された文字数までの文字列を返す.
const CString eseHankakuStrLeft(const CString& str, unsigned len);


// Donut.exeのフルパス名を返す
const CString GetExeFileName();

// exeのあるフォルダを返す(最後に \ がつく)
const CString GetExeDirectory();

// exeのあるフォルダを返す(最後に \ が付かない)
const CString GetExeDirName();


// UTF8文字列をwcs(UTF16LE)文字列に変換
const std::vector<wchar_t> utf8_to_wcs(const char* pUtf8);

const std::vector<TCHAR> utf8_to_tcs(const char* pUtf8);

const CString utf8_to_CString(const std::vector<char>& utf8);

// %20 等が使われたURLを通常の文字列に変換 ※に全角文字が混ざっていないこと！
const std::vector<char> urlstr_decode(const TCHAR* url);


// 文字列が半角英数記号のみでかつ%が混ざっている場合、UTF8をエンコードしたものとしてデコードする
const CString urlstr_decodeWhenASC(const CString& str);


/// 引数で指定したウィンドウを最前面表示する
void	SetForegroundWindow(HWND hWnd);


}	// namespace Misc

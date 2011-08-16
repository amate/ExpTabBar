/**
*	@file	ShellWrap.h
*	@brief	シェルに関する便利な関数
*/
#pragma once

#include <ShObjIdl.h>
#include <ShlObj.h>

namespace ShellWrap
{

/// アイテムＩＤリストからアイコンを作る
HICON	CreateIconFromIDList(PCIDLIST_ABSOLUTE pidl);


/// フルパスからアイテムＩＤリストを作成する
PIDLIST_ABSOLUTE CreateIDListFromFullPath(LPCTSTR strFullPath);

/// アイテムＩＤリストからフルパスを返す
CString	GetFullPathFromIDList(PCIDLIST_ABSOLUTE pidl);


/// アイテムＩＤリストの表示名を返す
CString GetNameFromIDList(PCIDLIST_ABSOLUTE pidl);


/// 現在エクスプローラーで表示中のフォルダのアイテムＩＤリストを作成する
PIDLIST_ABSOLUTE GetCurIDList(IShellBrowser* pShellBrowser);


/// アイテムＩＤリストで示されるフォルダが存在するかどうか
bool	IsExistFolderFromIDList(PCIDLIST_ABSOLUTE pidl);

LPBYTE	GetByteArrayFromBinary(int nSize, std::wstring strBinary);


}	// namespace ShellWrap

using namespace ShellWrap;








































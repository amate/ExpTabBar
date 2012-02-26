/**
*	@file	ShellWrap.h
*	@brief	�V�F���Ɋւ���֗��Ȋ֐�
*/
#pragma once

#include <ShObjIdl.h>
#include <ShlObj.h>

namespace ShellWrap
{

/// �A�C�e���h�c���X�g����A�C�R�������
HICON	CreateIconFromIDList(PCIDLIST_ABSOLUTE pidl);


/// �t���p�X����A�C�e���h�c���X�g���쐬����
PIDLIST_ABSOLUTE CreateIDListFromFullPath(LPCTSTR strFullPath);

/// �A�C�e���h�c���X�g����t���p�X��Ԃ�
CString	GetFullPathFromIDList(PCIDLIST_ABSOLUTE pidl);


/// �A�C�e���h�c���X�g�̕\������Ԃ�
CString GetNameFromIDList(PCIDLIST_ABSOLUTE pidl);


/// ���݃G�N�X�v���[���[�ŕ\�����̃t�H���_�̃A�C�e���h�c���X�g���쐬����
PIDLIST_ABSOLUTE GetCurIDList(IShellBrowser* pShellBrowser);

/// ���ݕ\�����̃r���[�� nIndex �ɂ���A�C�e����pidl�𓾂�
PIDLIST_ABSOLUTE GetIDListByIndex(IShellBrowser* pShellBrowser, int nIndex);

/// path ��InfoTipText���擾���܂�
CString GetInfoTipText(LPCTSTR path);

/// lnk�̃����N���Ԃ��܂�
PIDLIST_ABSOLUTE GetResolveIDList(PIDLIST_ABSOLUTE pidl);

/// .lnk�t�@�C�����쐬���܂�
bool	CreateLinkFile(LPITEMIDLIST pidl, LPCTSTR saveFilePath);

/// �A�C�e���h�c���X�g�Ŏ������t�H���_�����݂��邩�ǂ���
bool	IsExistFolderFromIDList(PCIDLIST_ABSOLUTE pidl);

LPBYTE	GetByteArrayFromBinary(int nSize, std::wstring strBinary);


}	// namespace ShellWrap

using namespace ShellWrap;








































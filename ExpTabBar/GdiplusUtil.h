/**
*	@file	GdiplusUtil.h
*	@brief	Gdi+���g���̂�֗��ɂ���
*/

#pragma once

#include <GdiPlus.h>

// ������/��n��
void	GdiplusInit();
void	GdiplusTerm();

//---------------------------------------
/// �g���q���w�肵�ăG���R�[�_�[���擾����
Gdiplus::ImageCodecInfo*	GetEncoderByExtension(LPCWSTR extension);

//--------------------------------------
/// MIME�^�C�v���w�肵�ăG���R�[�_���擾����
Gdiplus::ImageCodecInfo*	GetEncoderByMimeType(LPCWSTR mimetype);


















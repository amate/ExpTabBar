/**
*	@file	SharedMemoryUtil.h
*	@brief	�֗��ȋ��L�������쐬�N���X
*/

#pragma once

#include <windows.h>
#include <sstream>
#pragma warning( disable: 4244 )
#include <boost\archive\binary_iarchive.hpp>
#include <boost\archive\binary_oarchive.hpp>

/////////////////////////////////////////////
// CSharedMemoryT

template <bool t_bManaged>
class CSharedMemoryT
{
public:

	CSharedMemoryT() : m_hMap(NULL), m_pView(nullptr)
	{
	}

	~CSharedMemoryT()
	{
		if (t_bManaged)
			CloseHandle();
	}

	HANDLE	Handle() const { return m_hMap; }
	void*	GetPointer() const { return m_pView; }

	void	CloseHandle()
	{
		if (m_hMap) {
			if (m_pView) {
				::UnmapViewOfFile(m_pView);
				m_pView = nullptr;
			}
			::CloseHandle(m_hMap);
			m_hMap = NULL;
		}
	}

	void*	CreateSharedMemory(LPCTSTR sharedMemName, DWORD size)
	{
		ATLASSERT( m_hMap == NULL && m_pView == nullptr && sharedMemName );
		ATLASSERT( size >= 0 );
		m_hMap = ::CreateFileMapping(INVALID_HANDLE_VALUE, nullptr, PAGE_READWRITE, 0, size, sharedMemName);
		ATLASSERT( m_hMap );
		m_pView	= ::MapViewOfFile(m_hMap, FILE_MAP_ALL_ACCESS, 0, 0, 0);;
		ATLASSERT( m_pView );
		return m_pView;
	}

	void*	OpenSharedMemory(LPCTSTR sharedMemName, bool bReadOnly)
	{
		ATLASSERT( m_hMap == NULL && m_pView == nullptr && sharedMemName );
		DWORD dwDesiredAccess = bReadOnly ? FILE_MAP_READ : FILE_MAP_ALL_ACCESS;
		m_hMap = ::OpenFileMapping(dwDesiredAccess, FALSE, sharedMemName);
		if (m_hMap == NULL)
			return nullptr;
		ATLASSERT( m_hMap );
		m_pView	= ::MapViewOfFile(m_hMap, dwDesiredAccess, 0, 0, 0);;
		ATLASSERT( m_pView );
		return m_pView;
	}

	template <class T>
	bool	Serialize(const T& data, LPCTSTR sharedMemName = nullptr, bool bInheritHandle = false)
	{
		ATLASSERT( m_hMap == NULL );

		/* �V���A���C�Y */
		std::stringstream ss;
		boost::archive::binary_oarchive ar(ss);
		ar << data;
		std::string serializedData = ss.str();

		/* ���L�������쐬 */
		SECURITY_ATTRIBUTES	security_attributes = { sizeof(SECURITY_ATTRIBUTES) };
		security_attributes.bInheritHandle = bInheritHandle;
		DWORD	dwMapFileSize = serializedData.size() + sizeof(DWORD);
		m_hMap = ::CreateFileMapping(INVALID_HANDLE_VALUE, &security_attributes, PAGE_READWRITE, 0, dwMapFileSize, sharedMemName);
		ATLASSERT( m_hMap );
		if (m_hMap == NULL)
			return false;

		/* �������}�b�v���ꂽ�̈�ɏ������� */
		void* sharedMemData = (void*)::MapViewOfFile(m_hMap, FILE_MAP_ALL_ACCESS, 0, 0, 0);
		*(DWORD*)sharedMemData = dwMapFileSize;
		++((DWORD*&)sharedMemData);
		::memcpy_s(sharedMemData, dwMapFileSize, serializedData.c_str(), serializedData.size());
		--((DWORD*&)sharedMemData);
		::UnmapViewOfFile((LPCVOID)sharedMemData);
		return true;
	}

	template <class T>
	bool	Deserialize(T& data, LPCTSTR sharedMemName)
	{
		ATLASSERT( m_hMap == NULL );
		ATLASSERT( sharedMemName );

		/* ���L���������J�� */
		m_hMap = ::OpenFileMapping(FILE_MAP_READ, FALSE, sharedMemName);
		ATLASSERT( m_hMap );
		if (m_hMap == NULL)
			return false;

		_desirialize(data);
		return true;
	}

	template <class T>
	bool	Deserialize(T& data, HANDLE hMap)
	{
		ATLASSERT( m_hMap == NULL );
		ATLASSERT( hMap );

		m_hMap = hMap;

		_desirialize(data);
		return true;
	}

private:
	template <class T>
	void	_desirialize(T& data)
	{
		/* �������}�b�v���ꂽ�̈悩��f�[�^�����o�� */
		void* sharedMemData = (void*)::MapViewOfFile(m_hMap, FILE_MAP_READ, 0, 0, 0);
		DWORD dwSize = *(DWORD*)sharedMemData;
		++((DWORD*&)sharedMemData);
		std::stringstream ss;
		ss.write((const char*)sharedMemData, dwSize);
		--((DWORD*&)sharedMemData);
		::UnmapViewOfFile((LPCVOID)sharedMemData);

		/* �f�V���A���C�Y */
		boost::archive::binary_iarchive ar(ss);
		ar >> data;
	}

	HANDLE	m_hMap;
	void*	m_pView;
};


typedef CSharedMemoryT<false>	CSharedMemoryHandle;
typedef CSharedMemoryT<true>	CSharedMemory;



class CMutex
{
public:
	HANDLE m_hMutex;

	CMutex() : m_hMutex(NULL)
	{
	}

	~CMutex()
	{
		if (m_hMutex != NULL) {
			::CloseHandle(m_hMutex);
		}
	}

	bool	CreateMutex(LPCWSTR name)
	{
		ATLASSERT(m_hMutex == NULL);
		m_hMutex = ::CreateMutexW(nullptr, FALSE, name);
		if (m_hMutex == NULL) {
			DWORD error = GetLastError();
			return false;

		} else {
			return true;
		}
	}

	bool	OpenMutex(LPCWSTR name)
	{
		ATLASSERT(m_hMutex == NULL);
		m_hMutex = ::OpenMutexW(MUTEX_ALL_ACCESS, FALSE, name);
		if (m_hMutex == NULL) {
			DWORD error = GetLastError();
			return false;

		} else {
			return true;
		}
	}
};


class CMutexLock
{
public:
	CMutexLock(CMutex& mutex) : m_mutex(mutex)
	{
		DWORD ret = ::WaitForSingleObject(mutex.m_hMutex, INFINITE);
		if (ret != WAIT_OBJECT_0) {
			throw std::runtime_error("CMutexLock - WaitForSingleObject failed");
		}
	}

	~CMutexLock()
	{
		::ReleaseMutex(m_mutex.m_hMutex);
	}

private:
	CMutex& m_mutex;
};


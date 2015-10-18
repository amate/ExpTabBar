#pragma once

#include <Windows.h>
#include <string>
#include <vector>
#include <boost\serialization\serialization.hpp>
#include <boost\serialization\string.hpp>
#include <boost\serialization\vector.hpp>
#include <boost\archive\text_iarchive.hpp>
#include <boost\archive\text_oarchive.hpp>
#include <sstream>
#include <ShObjIdl.h>

struct OpenFolderAndSelectItems
{
	std::string pidlFolder;
	UINT cidl;
	std::vector<std::string>  apidl;
	DWORD dwFlags;

	OpenFolderAndSelectItems() : cidl(0), dwFlags(0) {}

	std::string	Serialize() const
	{
		std::stringstream ss;
		boost::archive::text_oarchive oa(ss);
		oa << (const OpenFolderAndSelectItems&)(*this);
		return ss.str();
		return "";
	}

	void	Deserialize(const std::string& data)
	{
		std::stringstream ss(data);
		boost::archive::text_iarchive ia(ss);
		ia >> (*this);
	}

	void DoOpenFolderAndSelectItems()
	{
		LPCITEMIDLIST* arrChildPidl = new LPCITEMIDLIST[cidl];
		for (UINT i = 0; i < cidl; ++i) {
			arrChildPidl[i] = (LPCITEMIDLIST)apidl[i].data();
		}
		::CoInitialize(NULL);
		HRESULT hr = ::SHOpenFolderAndSelectItems((LPCITEMIDLIST)pidlFolder.data(), cidl, arrChildPidl, dwFlags);
		//assert(hr == S_OK);
		::CoUninitialize();

		delete[] arrChildPidl;
	}

private:
	friend class boost::serialization::access;
	template<class Archive>
	void serialize(Archive& ar, const unsigned int ver)
	{
		ar & pidlFolder;
		ar & cidl;
		ar & apidl;
		ar & dwFlags;
	}
};

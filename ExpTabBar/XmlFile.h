// XmlFile.h
#pragma once

#include <XmlLite.h>

////////////////////////////////////////////////////////////////////////////
// CXmlFileRead

class CXmlFileRead
{
private:
	CComPtr<IXmlReader>	m_pReader;

public:
	// �R���X�g���N�^/�f�X�g���N�^
	CXmlFileRead(const CString &FilePath);
	~CXmlFileRead();

	HRESULT	Read(XmlNodeType* pNodeType) { return m_pReader->Read(pNodeType); }

	CString	GetLocalName();
	CString	GetValue();

	bool	GetInternalElement(LPWSTR externalElement, CString &strElement);

	bool	MoveToFirstAttribute() { return m_pReader->MoveToFirstAttribute() == S_OK; }
	bool	MoveToNextAttribute() { return m_pReader->MoveToNextAttribute() == S_OK; }
};


////////////////////////////////////////////////////////////////////////////
// CXmlFileWrite

class CXmlFileWrite
{
private:
	CComPtr<IXmlWriter> m_pWriter;

public:
	// �R���X�g���N�^/�f�X�g���N�^
	CXmlFileWrite(const CString &FilePath);
	~CXmlFileWrite();

	void	WriteStartElement(LPCWSTR Element);
	void	WriteElementString(LPCWSTR LocalName, LPCWSTR Value);
	void	WriteAttributeString(LPCWSTR LocalName, LPCWSTR Value);
	void	WriteAttributeValue(LPCWSTR LocalName, int nValue);
	void	WriteString(LPCWSTR	Text);
	void	WriteFullEndElement();
};



















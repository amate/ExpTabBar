// XmlFile.cpp

#include "stdafx.h"
#include "XmlFile.h"

///////////////////////////////////////////////////////////////////////////////
// CXmlFileWrite

// �R���X�g���N�^
CXmlFileRead::CXmlFileRead(const CString &FilePath)
{
	if (FilePath.IsEmpty()) {
		throw _T("�ۑ���̃p�X������܂���");
	}

    if (FAILED(CreateXmlReader(IID_IXmlReader, reinterpret_cast<void**>(&m_pReader), 0))){
        throw _T("CreateXmlReader���s");
    }
	
	// �t�@�C���X�g���[���쐬
    CComPtr<IStream> pStream;
    if (FAILED(SHCreateStreamOnFile(FilePath, STGM_READ, &pStream))){
        throw _T("SHCreateStreamOnFile���s");
    }

    if (FAILED(m_pReader->SetInput(pStream))){
        throw _T("SetInput���s");
    }

}

// �f�X�g���N�^
CXmlFileRead::~CXmlFileRead()
{
	//m_pReader->
}

// node == XmlNodeType_Element�̂Ƃ�element��Ԃ�
// MoveToFirstAttribute()����Α�������Ԃ�
CString	CXmlFileRead::GetLocalName()
{
	LPCWSTR	strLocalName;
	if (FAILED(m_pReader->GetLocalName(&strLocalName, NULL))) {
		throw _T("GetLocalName���s");
	}
	return CString(strLocalName);
}

CString	CXmlFileRead::GetValue()
{
	LPCWSTR	strValue;
	if (FAILED(m_pReader->GetValue(&strValue, NULL))) {
		throw _T("GetValue���s");
	}
	return CString(strValue);
}



// externalElement�̓����̗v�f��strElement�ŕԂ�
bool	CXmlFileRead::GetInternalElement(LPWSTR externalElement, CString &strElement)
{
	XmlNodeType 	nodeType;
	while (m_pReader->Read(&nodeType) == S_OK) {
		if (nodeType == XmlNodeType_Element) {
			strElement = GetLocalName();
			return true;
		} else if (nodeType == XmlNodeType_EndElement) {
			if (GetLocalName() == externalElement) {
				return false;	// </externalElement>�ɂ����̂�
			}
		}
	}
	throw _T("EnumElement���s");
	return false;
}








///////////////////////////////////////////////////////////////////////////////
// CXmlFileWrite

// �R���X�g���N�^
CXmlFileWrite::CXmlFileWrite(const CString &FilePath)
{
	if (FilePath.IsEmpty()) {
		throw _T("�ۑ���̃p�X������܂���");
	}

    if (FAILED(CreateXmlWriter(IID_IXmlWriter, reinterpret_cast<void**>(&m_pWriter), 0))){
		throw _T("CreateXmlWriter���s");
    }
   
	// �t�@�C���X�g���[���쐬
    CComPtr<IStream> pStream;
    if (FAILED(SHCreateStreamOnFile(FilePath, STGM_CREATE | STGM_WRITE, &pStream))){
        throw _T("SHCreateStreamOnFile���s");
    }

    if (FAILED(m_pWriter->SetOutput(pStream))){
        throw _T("SetOutput���s");
    }

    // �C���f���g�L����
    if (FAILED(m_pWriter->SetProperty(XmlWriterProperty_Indent, TRUE))){
        throw _T("SetProperty���s");
    }

    // <?xml version="1.0" encoding="UTF-8"?>
    if (FAILED(m_pWriter->WriteStartDocument(XmlStandalone_Omit))){
        throw _T("WriteStartDocument���s");
    }
}

// �f�X�g���N�^
CXmlFileWrite::~CXmlFileWrite()
{
	try {
		// �J���Ă���^�O�����
		if (FAILED(m_pWriter->WriteEndDocument())) {
			throw _T("WriteEndDocument���s");
		}

		// �t�@�C���ɏ�������
		if (FAILED(m_pWriter->Flush())) {
			throw _T("Flush���s");
		}
	}
	catch (LPCTSTR err) {
		ATLTRACE(err);
		err;
		ATLASSERT(FALSE);
	}
}




//	<Element>
void	CXmlFileWrite::WriteStartElement(LPCWSTR Element)
{
    if ( FAILED(m_pWriter->WriteStartElement(NULL, Element, NULL))){
        throw _T("WriteStartElement���s");
    }
}

//	<LocalName>Value</LocalName>
void	CXmlFileWrite::WriteElementString(LPCWSTR LocalName, LPCWSTR Value)
{
	if (FAILED(m_pWriter->WriteElementString(NULL, LocalName, NULL, Value))){
        throw _T("WriteElementString���s");
    }
}

//	<element	LocalName="Value">
void	CXmlFileWrite::WriteAttributeString(LPCWSTR LocalName, LPCWSTR Value)
{
	if (FAILED(m_pWriter->WriteAttributeString(NULL, LocalName, NULL, Value))){
        throw _T("WriteAttributeString���s");
    }
}

void	CXmlFileWrite::WriteAttributeValue(LPCWSTR LocalName, int nValue)
{
	CString str;
	str.Format(_T("%d"), nValue);
	WriteAttributeString(LocalName, str);
}

//	<element>	Text	</element>
void	CXmlFileWrite::WriteString(LPCWSTR	Text)
{
    if (FAILED(m_pWriter->WriteString(Text))){
        throw _T("WriteString���s");
    }
}

//	</Element>
void	CXmlFileWrite::WriteFullEndElement()
{
    if (FAILED(m_pWriter->WriteFullEndElement())){
        throw _T("WriteFullEndElement���s");
    }
}







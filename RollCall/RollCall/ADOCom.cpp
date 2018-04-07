#include "stdafx.h"
#include "ADOCom.h"


CADOCom::CADOCom()
{
}


CADOCom::~CADOCom()
{
}

void CADOCom::OnInitADOConn()
{
	//初始化OLE/COM库环境
	::CoInitialize(NULL);
	try
	{
		m_pConnection.CreateInstance("ADODB.Connection");
		_bstr_t csConnect = _T("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = .//NameData//NameData.accdb; Persist Security Info = False;");
		m_pConnection->Open(csConnect, "", "", adModeUnknown);
	}
	catch (_com_error e)
	{
		CString csTmp;
		csTmp.Format(_T("%X"), e.Error());
		AfxMessageBox(csTmp);
		AfxMessageBox(e.Description());
		AfxMessageBox(e.ErrorMessage());
	}
}


void CADOCom::ExitConnect()
{
	if (m_pRecordset != NULL)
		m_pRecordset->Close();
	m_pConnection->Close();
}

BOOL CADOCom::ExecuteSQL(_bstr_t bstrSQL)
{
	try
	{
		if (m_pConnection == NULL)
		{
			OnInitADOConn();
		}
		m_pConnection->Execute(bstrSQL, NULL, adCmdText);
		return TRUE;
	}
	catch (_com_error e)
	{
		CString csTmp;
		csTmp.Format(_T("%X"), e.Error());
		AfxMessageBox(csTmp);
		AfxMessageBox(e.Description());
		AfxMessageBox(e.ErrorMessage());
	}
}

_RecordsetPtr& CADOCom::GetRecordSet(_bstr_t bstrSQL)
{
	try
	{
		if (m_pConnection == NULL)
			OnInitADOConn();
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		m_pRecordset->Open(bstrSQL, m_pConnection.GetInterfacePtr(), adOpenDynamic, adLockOptimistic, adCmdText);
	}
	catch (_com_error e)
	{
		CString csTmp;
		csTmp.Format(_T("%X"), e.Error());
		AfxMessageBox(csTmp);
		AfxMessageBox(e.Description());
		AfxMessageBox(e.ErrorMessage());
	}
}


void CADOCom::CloseRecordset()
{
	if (m_pRecordset->GetState() == adStateOpen)
	{
		m_pRecordset->Close();
	}
}
#pragma once
class CADOCom
{
public:
	//ָ��connection�����ָ��
	_ConnectionPtr m_pConnection;
	//ָ��recordset �����ָ��
	_RecordsetPtr m_pRecordset;
public:
	CADOCom();
	virtual ~CADOCom();
	//��ʼ���������ݿ�
	void OnInitADOConn();
	//ִ�в�ѯ
	_RecordsetPtr& GetRecordSet(_bstr_t bstrSQL);
	void CADOCom::CloseRecordset();
	//ִ��SQL���
	BOOL ExecuteSQL(_bstr_t bstrSQL);
	//�Ͽ����ݿ�����
	void ExitConnect();
};


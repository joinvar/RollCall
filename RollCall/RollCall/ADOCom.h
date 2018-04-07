#pragma once
class CADOCom
{
public:
	//指向connection对象的指针
	_ConnectionPtr m_pConnection;
	//指向recordset 对象的指针
	_RecordsetPtr m_pRecordset;
public:
	CADOCom();
	virtual ~CADOCom();
	//初始化连接数据库
	void OnInitADOConn();
	//执行查询
	_RecordsetPtr& GetRecordSet(_bstr_t bstrSQL);
	void CADOCom::CloseRecordset();
	//执行SQL语句
	BOOL ExecuteSQL(_bstr_t bstrSQL);
	//断开数据库连接
	void ExitConnect();
};


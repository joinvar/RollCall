
// RollCallDlg.h : ͷ�ļ�
//

#pragma once

#include "ADOCom.h"
// CRollCallDlg �Ի���
class CRollCallDlg : public CDialogEx
{
// ����
public:
	CRollCallDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_ROLLCALL_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()



public:
	CFont m_Font;   //����
	CADOCom m_AdoCon;//ADO ����
	BOOL	m_bIsUpData;

public:
	afx_msg HBRUSH OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor);
	afx_msg void OnBtnNextName();
};

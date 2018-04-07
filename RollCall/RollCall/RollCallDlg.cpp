
// RollCallDlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "RollCall.h"
#include "RollCallDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CRollCallDlg �Ի���



CRollCallDlg::CRollCallDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CRollCallDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CRollCallDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CRollCallDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_WM_CTLCOLOR()
	ON_BN_CLICKED(IDC_BUTTON1, &CRollCallDlg::OnBtnNextName)
END_MESSAGE_MAP()


// CRollCallDlg ��Ϣ�������

BOOL CRollCallDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// ���������С����ʽ
	m_Font.CreatePointFont(560, _T("����"));
	GetDlgItem(IDC_STATICNAME)->SetFont(&m_Font);
	m_AdoCon.OnInitADOConn();	
	m_bIsUpData = FALSE;//Ĭ�ϵ�һ�θ���
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}

void CRollCallDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CRollCallDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CRollCallDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



HBRUSH CRollCallDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  �ڴ˸��� DC ���κ�����
	if (pWnd->GetDlgCtrlID() == IDC_STATICNAME)
	{
		pDC->SetTextColor(RGB(0, 0, 255));
	}
	// TODO:  ���Ĭ�ϵĲ������軭�ʣ��򷵻���һ������
	return hbr;
}


void CRollCallDlg::OnBtnNextName()
{
	SetDlgItemText(IDC_BUTTON1, _T("��һ��"));
	try
	{
		CString csSQL;
		m_AdoCon.m_pRecordset.CreateInstance("ADODB.Recordset");
		//��UsedName �ֶ��ÿ�
		if (m_bIsUpData == FALSE)
		{
			m_bIsUpData = TRUE;
			csSQL = _T("select * from AllName");			
			m_AdoCon.m_pRecordset->Open((_bstr_t)(_variant_t)csSQL, m_AdoCon.m_pConnection.GetInterfacePtr(),
				adOpenKeyset, adLockOptimistic, -1);
			while (m_AdoCon.m_pRecordset->adoEOF == 0)
			{
				m_AdoCon.m_pRecordset->PutCollect("UsedName", (_bstr_t)"");
				m_AdoCon.m_pRecordset->MoveNext();
			}
			if (m_AdoCon.m_pRecordset != NULL)
			{
				m_AdoCon.m_pRecordset->Close();
			}
		}
		
		//�����ݿ��������ȡһ�����������ļ�¼
		csSQL = _T("select * from AllName where UsedName = '' order by right(cstr(rnd(-int(rnd(-timer())*100+id)))*1000*Now(),2) "); //order by rnd(ID)"
		CString csValue = _T("");		
		m_AdoCon.m_pRecordset->Open((_bstr_t)(_variant_t)csSQL, m_AdoCon.m_pConnection.GetInterfacePtr(), adOpenKeyset, adLockOptimistic, -1);
		if (!m_AdoCon.m_pRecordset->adoEOF)
		{
			//��ȡ����
			csValue = (char*)(_bstr_t)m_AdoCon.m_pRecordset->GetCollect("AllName");
			//��ȡ����UsedName,��ֹ�´λ�ȡ�����ظ�
			//m_AdoCon.m_pRecordset->MoveFirst();  //�����¾䶼������
			m_AdoCon.m_pRecordset->Move((long)0, vtMissing);
			m_AdoCon.m_pRecordset->PutCollect("UsedName", (_bstr_t)csValue);
			m_AdoCon.m_pRecordset->Update();
			SetDlgItemText(IDC_STATICNAME, csValue);
		}
		else
		{
			SetDlgItemText(IDC_STATICNAME, _T("�������"));
		}		
		if (m_AdoCon.m_pRecordset != NULL)
		{
			m_AdoCon.m_pRecordset->Close();
		}

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


// RollCallDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "RollCall.h"
#include "RollCallDlg.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CRollCallDlg 对话框



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


// CRollCallDlg 消息处理程序

BOOL CRollCallDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// 设置字体大小和样式
	m_Font.CreatePointFont(560, _T("宋体"));
	GetDlgItem(IDC_STATICNAME)->SetFont(&m_Font);
	m_AdoCon.OnInitADOConn();	
	m_bIsUpData = FALSE;//默认第一次更新
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CRollCallDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CRollCallDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



HBRUSH CRollCallDlg::OnCtlColor(CDC* pDC, CWnd* pWnd, UINT nCtlColor)
{
	HBRUSH hbr = CDialogEx::OnCtlColor(pDC, pWnd, nCtlColor);

	// TODO:  在此更改 DC 的任何特性
	if (pWnd->GetDlgCtrlID() == IDC_STATICNAME)
	{
		pDC->SetTextColor(RGB(0, 0, 255));
	}
	// TODO:  如果默认的不是所需画笔，则返回另一个画笔
	return hbr;
}


void CRollCallDlg::OnBtnNextName()
{
	SetDlgItemText(IDC_BUTTON1, _T("下一个"));
	try
	{
		CString csSQL;
		m_AdoCon.m_pRecordset.CreateInstance("ADODB.Recordset");
		//将UsedName 字段置空
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
		
		//从数据库中随机获取一条符合条件的记录
		csSQL = _T("select * from AllName where UsedName = '' order by right(cstr(rnd(-int(rnd(-timer())*100+id)))*1000*Now(),2) "); //order by rnd(ID)"
		CString csValue = _T("");		
		m_AdoCon.m_pRecordset->Open((_bstr_t)(_variant_t)csSQL, m_AdoCon.m_pConnection.GetInterfacePtr(), adOpenKeyset, adLockOptimistic, -1);
		if (!m_AdoCon.m_pRecordset->adoEOF)
		{
			//获取名字
			csValue = (char*)(_bstr_t)m_AdoCon.m_pRecordset->GetCollect("AllName");
			//获取后置UsedName,防止下次获取名字重复
			//m_AdoCon.m_pRecordset->MoveFirst();  //这句和下句都可以用
			m_AdoCon.m_pRecordset->Move((long)0, vtMissing);
			m_AdoCon.m_pRecordset->PutCollect("UsedName", (_bstr_t)csValue);
			m_AdoCon.m_pRecordset->Update();
			SetDlgItemText(IDC_STATICNAME, csValue);
		}
		else
		{
			SetDlgItemText(IDC_STATICNAME, _T("点名完毕"));
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

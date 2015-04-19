// ADOSQLDlg.cpp : implementation file
//

#include "stdafx.h"
#include "ADOSQL.h"
#include "ADOSQLDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif
extern CADOSQLApp theApp;
/////////////////////////////////////////////////////////////////////////////
// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();
	
	// Dialog Data
	//{{AFX_DATA(CAboutDlg)
	enum { IDD = IDD_ABOUTBOX };
	//}}AFX_DATA
	
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CAboutDlg)
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	//}}AFX_VIRTUAL
	
	// Implementation
protected:
	//{{AFX_MSG(CAboutDlg)
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
	//{{AFX_DATA_INIT(CAboutDlg)
	//}}AFX_DATA_INIT
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CAboutDlg)
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
//{{AFX_MSG_MAP(CAboutDlg)
// No message handlers
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CADOSQLDlg dialog

CADOSQLDlg::CADOSQLDlg(CWnd* pParent /*=NULL*/)
: CDialog(CADOSQLDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CADOSQLDlg)
	m_sno = _T("");
	m_age = _T("");
	m_dept = _T("");
	m_comment = _T("");
	m_sname = _T("");
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CADOSQLDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CADOSQLDlg)
	DDX_Control(pDX, IDC_LIST1, m_list);
	DDX_Text(pDX, IDC_EDIT1, m_sno);
	DDX_Text(pDX, IDC_EDIT3, m_age);
	DDX_Text(pDX, IDC_EDIT4, m_dept);
	DDX_Text(pDX, IDC_EDIT5, m_comment);
	DDX_Text(pDX, IDC_EDIT2, m_sname);
	//}}AFX_DATA_MAP
}

BEGIN_MESSAGE_MAP(CADOSQLDlg, CDialog)
//{{AFX_MSG_MAP(CADOSQLDlg)
ON_WM_SYSCOMMAND()
ON_WM_PAINT()
ON_WM_QUERYDRAGICON()
ON_BN_CLICKED(IDC_BUTTON1, OnInsert)
ON_BN_CLICKED(IDC_BUTTON2, OnChange)
ON_BN_CLICKED(IDC_BUTTON3, OnDelete)
ON_BN_CLICKED(IDC_BUTTON4, OnFind)
ON_BN_CLICKED(IDC_BUTTON5, OnFresh)
ON_LBN_SELCHANGE(IDC_LIST1, OnSelchangeList1)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CADOSQLDlg message handlers

BOOL CADOSQLDlg::OnInitDialog()
{
	
	CDialog::OnInitDialog();
	
	
	HRESULT hr;  
	try  
	{  
		hr = m_pConnection.CreateInstance("ADODB.Connection");
		if(SUCCEEDED(hr))  
		{  
			hr = m_pConnection->Open("driver={SQL Server};Server=8BCB;DATABASE=Student;UID=sa;PWD=123","","",adModeUnknown);
		}
	}
	catch(_com_error e)
	{
		AfxMessageBox("shibai!");
	}
	m_pRecordset.CreateInstance(__uuidof(Recordset));
	try
	{
		m_pRecordset->Open("SELECT * FROM s", 
			m_pConnection.GetInterfacePtr(), 
			adOpenDynamic,
			adLockOptimistic,
			adCmdText);
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
	
	ListData();
	
	// Add "About..." menu item to system menu.
	
	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);
	
	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}
	
	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CADOSQLDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CADOSQLDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting
		
		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);
		
		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;
		
		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CADOSQLDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CADOSQLDlg::OnInsert() 
{
	// TODO: Add your control notification handler code here
	UpdateData();
	if(m_sname == "" || m_sno == "")
	{
		AfxMessageBox("学号和姓名不能为空!");
		return;
	}
	
	try
	{
		
		m_pRecordset->AddNew();
		m_pRecordset->PutCollect("sno", _variant_t(m_sno));
		m_pRecordset->PutCollect("sname", _variant_t(m_sname));
		m_pRecordset->PutCollect("age", atol(m_age));
		m_pRecordset->PutCollect("dept", _variant_t(m_dept));
		m_pRecordset->PutCollect("comment", _variant_t(m_comment));
		m_pRecordset->Update();
		AfxMessageBox("插入成功!");
		
		int nCurSel = m_list.GetCurSel();
		ListData();
		m_list.SetCurSel(nCurSel);
		
		OnSelchangeList1();
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
}

void CADOSQLDlg::OnChange() 
{
	// TODO: Add your control notification handler code here
	UpdateData();
	if(m_list.GetCount() == 0)
	{
		AfxMessageBox("???????!");
		return;
	}
	else if(m_list.GetCurSel() < 0 || m_list.GetCurSel() > m_list.GetCount())
		m_list.SetCurSel(0);
	
	try
	{
		m_pRecordset->PutCollect("sno", _variant_t(m_sno));
		m_pRecordset->PutCollect("sname", _variant_t(m_sname));
		m_pRecordset->PutCollect("age", atol(m_age));
		m_pRecordset->PutCollect("dept", _variant_t(m_dept));
		m_pRecordset->PutCollect("comment", _variant_t(m_comment));
		m_pRecordset->Update();
		
		int nCurSel = m_list.GetCurSel();
		ListData();
		m_list.SetCurSel(nCurSel);
		
		OnSelchangeList1();
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
}

void CADOSQLDlg::OnDelete() 
{
	// TODO: Add your control notification handler code here
	if(m_list.GetCount() == 0)
		return;
	else if(m_list.GetCurSel() < 0 || m_list.GetCurSel() > m_list.GetCount())
		m_list.SetCurSel(0);
	try
	{
		
		m_pRecordset->Delete(adAffectCurrent);
		m_pRecordset->Update();
		
		int nCurSel = m_list.GetCurSel();
		m_list.DeleteString(nCurSel);
		if(nCurSel == 0 && (m_list.GetCount() != 0))
			m_list.SetCurSel(nCurSel);
		else if(m_list.GetCount() != 0)
			m_list.SetCurSel(nCurSel-1);
		
		OnSelchangeList1();
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
}

void CADOSQLDlg::OnFind() 
{
	// TODO: Add your control notification handler code here
	UpdateData(TRUE);
	
	CString strSql;
	if(m_sname != "" && m_sno != "")
		strSql.Format("SELECT * FROM s WHERE sname = '%s' AND sno = '%s'",
		m_sname,m_sno);
	else if(m_sname != "" && m_sno == "")
		strSql.Format("SELECT * FROM s WHERE sname = '%s'",m_sname);
	else if(m_sname == "" && m_sno != "")
		strSql.Format("SELECT * FROM s WHERE sno = '%s'",m_sno);
	else
		strSql = "SELECT * FROM s";
	try
	{
		
		m_pRecordset->Close();
		
		m_pRecordset->Open(strSql.AllocSysString(),                
			theApp.m_pConnection.GetInterfacePtr(),
			adOpenDynamic,
			adLockOptimistic,
			adCmdText);
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
	
	ListData();
}

void CADOSQLDlg::OnFresh() 
{
	// TODO: Add your control notification handler code here
	m_sno="";
	m_sname="";
	m_age="";
	m_dept="";
	m_comment="";
	m_list.ResetContent();
	UpdateData(false);
}

void CADOSQLDlg::OnSelchangeList1() 
{
	// TODO: Add your control notification handler code here
	int curSel = m_list.GetCurSel();
	_variant_t var,varIndex;
	if(curSel < 0)
		return;
	try
	{
		m_pRecordset->MoveFirst();
		m_pRecordset->Move(long(curSel));
		var = m_pRecordset->GetCollect("sname");
		if(var.vt != VT_NULL)
			m_sname = (LPCSTR)_bstr_t(var);
		var = m_pRecordset->GetCollect("age");
		if(var.vt != VT_NULL)
			m_age=(LPCSTR)_bstr_t(var);
		var = m_pRecordset->GetCollect("sno");
		if(var.vt != VT_NULL)
			m_sno = (LPCSTR)_bstr_t(var);
		var = m_pRecordset->GetCollect("dept");
		if(var.vt != VT_NULL)
			m_dept = (LPCSTR)_bstr_t(var);
		var = m_pRecordset->GetCollect("comment");
		if(var.vt != VT_NULL)
			m_comment = (LPCSTR)_bstr_t(var);
		UpdateData(false);
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
}

void CADOSQLDlg::ListData()
{
	_variant_t var;
	CString strname,strage,strsno,strdept,strcom;
	m_list.ResetContent();
	strname=strage=strsno=strdept=strcom="";
	try
	{
		if(!m_pRecordset->BOF)
			m_pRecordset->MoveFirst();
		else
		{
			AfxMessageBox("??????");
			return;
		}
		while(!m_pRecordset->adoEOF)
		{
			var = m_pRecordset->GetCollect("sno");
			if(var.vt != VT_NULL)
				strsno = (LPCSTR)_bstr_t(var);
			var = m_pRecordset->GetCollect("sname");
			if(var.vt != VT_NULL)
				strname = (LPCSTR)_bstr_t(var);
			var = m_pRecordset->GetCollect("age");
			if(var.vt != VT_NULL)
				strage = (LPCSTR)_bstr_t(var);
			var = m_pRecordset->GetCollect("dept");
			if(var.vt != VT_NULL)
				strdept = (LPCSTR)_bstr_t(var);
			var = m_pRecordset->GetCollect("comment");
			if(var.vt != VT_NULL)
				strcom = (LPCSTR)_bstr_t(var);
			m_list.AddString( strsno + " --> " +strname + " --> "+strage+ " --> "+ strdept+ " --> "+ strcom);
			m_pRecordset->MoveNext();
		}
		m_list.SetCurSel(0);
		OnSelchangeList1();
	}
	catch(_com_error *e)
	{
		AfxMessageBox(e->ErrorMessage());
	}
}

// ADOSQLDlg.h : header file
//

#if !defined(AFX_ADOSQLDLG_H__2284CE03_06EF_4BD0_861A_4E4DC3BE8968__INCLUDED_)
#define AFX_ADOSQLDLG_H__2284CE03_06EF_4BD0_861A_4E4DC3BE8968__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CADOSQLDlg dialog

class CADOSQLDlg : public CDialog
{
// Construction
public:
	void ListData();
	CADOSQLDlg(CWnd* pParent = NULL);	// standard constructor


	_ConnectionPtr m_pConnection;
	_CommandPtr m_pCommand;
	_RecordsetPtr m_pRecordset;

// Dialog Data
	//{{AFX_DATA(CADOSQLDlg)
	enum { IDD = IDD_ADOSQL_DIALOG };
	CListBox	m_list;
	CString	m_sno;
	CString	m_age;
	CString	m_dept;
	CString	m_comment;
	CString	m_sname;
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CADOSQLDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CADOSQLDlg)
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnInsert();
	afx_msg void OnChange();
	afx_msg void OnDelete();
	afx_msg void OnFind();
	afx_msg void OnFresh();
	afx_msg void OnSelchangeList1();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADOSQLDLG_H__2284CE03_06EF_4BD0_861A_4E4DC3BE8968__INCLUDED_)

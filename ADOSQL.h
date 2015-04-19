// ADOSQL.h : main header file for the ADOSQL application
//

#if !defined(AFX_ADOSQL_H__D9B85338_2B6E_45AF_9E38_6125D249AE6C__INCLUDED_)
#define AFX_ADOSQL_H__D9B85338_2B6E_45AF_9E38_6125D249AE6C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CADOSQLApp:
// See ADOSQL.cpp for the implementation of this class
//

class CADOSQLApp : public CWinApp
{
public:
	_ConnectionPtr m_pConnection;
_CommandPtr m_pCommand;
_RecordsetPtr     m_pRecordset;

	CADOSQLApp();
	
// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CADOSQLApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CADOSQLApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_ADOSQL_H__D9B85338_2B6E_45AF_9E38_6125D249AE6C__INCLUDED_)

#if !defined(AFX_GTF_PASSPORT_COMM_ACTICEXPPG_H__3E6D7448_0DED_4457_A795_3B39C4200F71__INCLUDED_)
#define AFX_GTF_PASSPORT_COMM_ACTICEXPPG_H__3E6D7448_0DED_4457_A795_3B39C4200F71__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// GTF_PASSPORT_COMM_ActiceXPpg.h : Declaration of the CGTF_PASSPORT_COMM_ActiceXPropPage property page class.

////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXPropPage : See GTF_PASSPORT_COMM_ActiceXPpg.cpp.cpp for implementation.

class CGTF_PASSPORT_COMM_ActiceXPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CGTF_PASSPORT_COMM_ActiceXPropPage)
	DECLARE_OLECREATE_EX(CGTF_PASSPORT_COMM_ActiceXPropPage)

// Constructor
public:
	CGTF_PASSPORT_COMM_ActiceXPropPage();

// Dialog Data
	//{{AFX_DATA(CGTF_PASSPORT_COMM_ActiceXPropPage)
	enum { IDD = IDD_PROPPAGE_GTF_PASSPORT_COMM_ACTICEX };
		// NOTE - ClassWizard will add data members here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	//{{AFX_MSG(CGTF_PASSPORT_COMM_ActiceXPropPage)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GTF_PASSPORT_COMM_ACTICEXPPG_H__3E6D7448_0DED_4457_A795_3B39C4200F71__INCLUDED)

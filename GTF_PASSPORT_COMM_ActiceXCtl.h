#if !defined(AFX_GTF_PASSPORT_COMM_ACTICEXCTL_H__00C249FE_04B7_436E_85EF_ADB6352F719A__INCLUDED_)
#define AFX_GTF_PASSPORT_COMM_ACTICEXCTL_H__00C249FE_04B7_436E_85EF_ADB6352F719A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// GTF_PASSPORT_COMM_ActiceXCtl.h : Declaration of the CGTF_PASSPORT_COMM_ActiceXCtrl ActiveX Control class.

/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXCtrl : See GTF_PASSPORT_COMM_ActiceXCtl.cpp for implementation.

class CGTF_PASSPORT_COMM_ActiceXCtrl : public COleControl
{
	DECLARE_DYNCREATE(CGTF_PASSPORT_COMM_ActiceXCtrl)

// Constructor
public:
	CGTF_PASSPORT_COMM_ActiceXCtrl();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CGTF_PASSPORT_COMM_ActiceXCtrl)
	public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	
	//}}AFX_VIRTUAL

// Implementation
protected:
	~CGTF_PASSPORT_COMM_ActiceXCtrl();

	DECLARE_OLECREATE_EX(CGTF_PASSPORT_COMM_ActiceXCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CGTF_PASSPORT_COMM_ActiceXCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CGTF_PASSPORT_COMM_ActiceXCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CGTF_PASSPORT_COMM_ActiceXCtrl)		// Type name and misc status

// Message maps
	//{{AFX_MSG(CGTF_PASSPORT_COMM_ActiceXCtrl)
		// NOTE - ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	//{{AFX_DISPATCH(CGTF_PASSPORT_COMM_ActiceXCtrl)
	afx_msg short OpenPort();
	afx_msg short ClosePort();
	afx_msg short Scan();
	afx_msg short ScanCancel();
	afx_msg short ReceiveData(short TimeOut);
	afx_msg BSTR GetMRZ1();
	afx_msg BSTR GetMRZ2();
	afx_msg BSTR GetPassportInfo();
	afx_msg short Clear();
	afx_msg short CheckReceiveData();
	afx_msg short OpenPortByNum(short PorNum);
	afx_msg short OpenPortByNumber(short PorNum);
	afx_msg BSTR testAlert();
	afx_msg short IsOpen();
	afx_msg short SetPassportScanType(short nType);
	//}}AFX_DISPATCH
	DECLARE_DISPATCH_MAP()

// Event maps
	//{{AFX_EVENT(CGTF_PASSPORT_COMM_ActiceXCtrl)
	//}}AFX_EVENT
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	
	enum {
	//{{AFX_DISP_ID(CGTF_PASSPORT_COMM_ActiceXCtrl)
	dispidOpenPort =		1L,
	dispidClosePort =		2L,
	dispidScan =			3L,
	dispidScanCancel =		4L,
	dispidReceiveData =		5L,
	dispidGetMRZ1 =			6L,
	dispidGetMRZ2 =			7L,
	dispidGetPassportInfo = 8L,
	dispidClear =			9L,
	dispidCheckReceiveData = 10L,
	dispidOpenPortByNum =	 11L,
	dispidOpenPortByNumber = 12L,
	dispidTestAlert =		13L,
	dispidIsOpen 	= 14L,
	dispidSetPassportScanType 	= 15L,
	//}}AFX_DISP_ID
	};
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GTF_PASSPORT_COMM_ACTICEXCTL_H__00C249FE_04B7_436E_85EF_ADB6352F719A__INCLUDED)

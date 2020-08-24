#if !defined(AFX_GTF_PASSPORT_COMM_ACTICEX_H__42AE3001_2455_44F6_91E8_4C3C5531381D__INCLUDED_)
#define AFX_GTF_PASSPORT_COMM_ACTICEX_H__42AE3001_2455_44F6_91E8_4C3C5531381D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// GTF_PASSPORT_COMM_ActiceX.h : main header file for GTF_PASSPORT_COMM_ACTICEX.DLL

#if !defined( __AFXCTL_H__ )
	#error include 'afxctl.h' before including this file
#endif

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXApp : See GTF_PASSPORT_COMM_ActiceX.cpp for implementation.

class CGTF_PASSPORT_COMM_ActiceXApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_GTF_PASSPORT_COMM_ACTICEX_H__42AE3001_2455_44F6_91E8_4C3C5531381D__INCLUDED)

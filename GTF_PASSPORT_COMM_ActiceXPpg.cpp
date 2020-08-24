// GTF_PASSPORT_COMM_ActiceXPpg.cpp : Implementation of the CGTF_PASSPORT_COMM_ActiceXPropPage property page class.

#include "stdafx.h"
#include "GTF_PASSPORT_COMM_ActiceX.h"
#include "GTF_PASSPORT_COMM_ActiceXPpg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


IMPLEMENT_DYNCREATE(CGTF_PASSPORT_COMM_ActiceXPropPage, COlePropertyPage)


/////////////////////////////////////////////////////////////////////////////
// Message map

BEGIN_MESSAGE_MAP(CGTF_PASSPORT_COMM_ActiceXPropPage, COlePropertyPage)
	//{{AFX_MSG_MAP(CGTF_PASSPORT_COMM_ActiceXPropPage)
	// NOTE - ClassWizard will add and remove message map entries
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()


/////////////////////////////////////////////////////////////////////////////
// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CGTF_PASSPORT_COMM_ActiceXPropPage, "GTFPASSPORTCOMMACTICEX.GTFPASSPORTCOMMActiceXPropPage.1",
	0xd87cb8f3, 0x3656, 0x4e62, 0x90, 0xf3, 0xf6, 0xcd, 0x6b, 0xe0, 0x9, 0x61)


/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXPropPage::CGTF_PASSPORT_COMM_ActiceXPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CGTF_PASSPORT_COMM_ActiceXPropPage

BOOL CGTF_PASSPORT_COMM_ActiceXPropPage::CGTF_PASSPORT_COMM_ActiceXPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_GTF_PASSPORT_COMM_ACTICEX_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}


/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXPropPage::CGTF_PASSPORT_COMM_ActiceXPropPage - Constructor

CGTF_PASSPORT_COMM_ActiceXPropPage::CGTF_PASSPORT_COMM_ActiceXPropPage() :
	COlePropertyPage(IDD, IDS_GTF_PASSPORT_COMM_ACTICEX_PPG_CAPTION)
{
	//{{AFX_DATA_INIT(CGTF_PASSPORT_COMM_ActiceXPropPage)
	// NOTE: ClassWizard will add member initialization here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA_INIT
}


/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXPropPage::DoDataExchange - Moves data between page and properties

void CGTF_PASSPORT_COMM_ActiceXPropPage::DoDataExchange(CDataExchange* pDX)
{
	//{{AFX_DATA_MAP(CGTF_PASSPORT_COMM_ActiceXPropPage)
	// NOTE: ClassWizard will add DDP, DDX, and DDV calls here
	//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_DATA_MAP
	DDP_PostProcessing(pDX);
}


/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXPropPage message handlers

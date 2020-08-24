// GTF_PASSPORT_COMM_ActiceX.cpp : Implementation of CGTF_PASSPORT_COMM_ActiceXApp and DLL registration.

#include "stdafx.h"
#include "GTF_PASSPORT_COMM_ActiceX.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif


CGTF_PASSPORT_COMM_ActiceXApp NEAR theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0x91863ad0, 0xc3a7, 0x4687, { 0xb1, 0xd2, 0xae, 0xf2, 0x57, 0x69, 0x71, 0x7e } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;


////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXApp::InitInstance - DLL initialization

BOOL CGTF_PASSPORT_COMM_ActiceXApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}


////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMM_ActiceXApp::ExitInstance - DLL termination

int CGTF_PASSPORT_COMM_ActiceXApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

	return COleControlModule::ExitInstance();
}


/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}


/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}

// GTF_PASSPORT_COMM.cpp : Defines the initialization routines for the DLL.
//

#include "stdafx.h"
#include "GTF_PASSPORT_COMM.h"
#include "ComDef.h"
#include "Comm.h"
#include "CommCtrl.h"

void HexDump(unsigned char *pDcs, int len);

// {	Serial Communication
// MRZ data size
#define DATA_LEN	106

extern CCommCtrl gCom;

extern INT AutoDetect(HWND hWnd, INT iBaudRate , int nType);

static UINT	g_nPortNum = 0;
static UINT	g_nBaudRate = 115200;
static BOOL g_bCommOpened = FALSE;

static BYTE	g_CommBuf[512];
// }


HWND g_hWnd;
CString g_MRZ1;
CString g_MRZ2;

CRITICAL_SECTION g_cx;

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

//
//	Note!
//
//		If this DLL is dynamically linked against the MFC
//		DLLs, any functions exported from this DLL which
//		call into MFC must have the AFX_MANAGE_STATE macro
//		added at the very beginning of the function.
//
//		For example:
//
//		extern "C" BOOL PASCAL EXPORT ExportedFunction()
//		{
//			AFX_MANAGE_STATE(AfxGetStaticModuleState());
//			// normal function body here
//		}
//
//		It is very important that this macro appear in each
//		function, prior to any calls into MFC.  This means that
//		it must appear as the first statement within the 
//		function, even before any object variable declarations
//		as their constructors may generate calls into the MFC
//		DLL.
//
//		Please see MFC Technical Notes 33 and 58 for additional
//		details.
//

/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMMApp

BEGIN_MESSAGE_MAP(CGTF_PASSPORT_COMMApp, CWinApp)
	//{{AFX_MSG_MAP(CGTF_PASSPORT_COMMApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CGTF_PASSPORT_COMMApp construction

CGTF_PASSPORT_COMMApp::CGTF_PASSPORT_COMMApp()
{
	// TODO: add construction code here,
	// Place all significant initialization in InitInstance
}

/////////////////////////////////////////////////////////////////////////////
// The one and only CGTF_PASSPORT_COMMApp object

CGTF_PASSPORT_COMMApp theApp;

extern "C" __declspec(dllexport) int __stdcall OpenPort();
extern "C" __declspec(dllexport) int __stdcall ClosePort();
extern "C" __declspec(dllexport) int __stdcall Scan();
extern "C" __declspec(dllexport) int __stdcall ScanCancel();
extern "C" __declspec(dllexport) int __stdcall ReceiveData(int TimeOut);
extern "C" __declspec(dllexport) int __stdcall GetMRZ1(char *refMRZ1);
extern "C" __declspec(dllexport) int __stdcall GetMRZ2(char *refMRZ2);
extern "C" __declspec(dllexport) int __stdcall GetPassportInfo(char *refPassInfo);
extern "C" __declspec(dllexport) int __stdcall Clear();
extern "C" __declspec(dllexport) int __stdcall CheckReceiveData();
extern "C" __declspec(dllexport) int __stdcall OpenPortByNumber(int PorNum);
extern "C" __declspec(dllexport) int __stdcall IsOpen();
{
	g_nPortNum = AutoDetect(g_hWnd, g_nBaudRate);
	//printf("port num : %d--->%d\n", g_nPortNum, g_bCommOpened);
	if( !g_bCommOpened ){
		g_bCommOpened = gCom.OpenPort(g_hWnd, g_nPortNum, g_nBaudRate );
	}else{
		gCom.SetHwnd(g_hWnd);
	}

	if( g_bCommOpened )
	{
		::InitializeCriticalSection(&g_cx);

		return 1;
	}

	return 0;
}


extern "C" __declspec(dllexport) int __stdcall OpenPortByNumber(int PorNum)
{
	if( !g_bCommOpened ){
		g_bCommOpened = gCom.OpenPort(g_hWnd, PorNum, g_nBaudRate );
	}else{
		gCom.SetHwnd(g_hWnd);
	}

	if( g_bCommOpened )
	{
		::InitializeCriticalSection(&g_cx);

		return 1;
	}

	return 0;
}


enum
{
	CHK_STX,
	CHK_LENGTH,
	CHK_COMMAND,
	CHK_DATA,
	CHK_ETX,
	CHK_LRC
};

int ExceptionReceiveData(){

    gCom.m_Queue.GetData( 200, g_CommBuf );

	return 1;
}


int ReceiveCommData(){

    int     cnt;

	char    MRZ1[45];
	char    MRZ2[45];
	int     rcv_len=0;
	int     state=CHK_STX;
	BYTE    rcv_packet[512];

    int     len=0;
	int run_f=1;

    cnt = gCom.m_Queue.GetData( DATA_LEN, g_CommBuf );
	printf("%d ", cnt);
	if(!cnt)
	{
		return 0;
	}
	else
	{
		memset(rcv_packet,	0x00,	sizeof(rcv_packet));
		memcpy((char*)&rcv_packet[rcv_len],(const char*)g_CommBuf, cnt);
	}

	HexDump((unsigned char*)rcv_packet,  cnt);	

	rcv_len+=cnt;

	while(run_f)
	{
		switch(state)
		{
			case CHK_STX:
				if(rcv_packet[0]==0x02)
				{
					state=CHK_LENGTH;
				}
				break;
			case CHK_LENGTH:
				if(rcv_len>2)
				{
					//len =  *( (unsigned short *)&rcv_packet[1]);
					len =  rcv_packet[1]<<8 | rcv_packet[2];
					if(rcv_len> (len+3))
					{
						state=CHK_COMMAND;
					}
					else
					{
						return 0;
					}
				}
				else
					return 0;

				break;
			case CHK_COMMAND:
				if(rcv_packet[3]==0xB1)
				{
					state=CHK_DATA;
				}
				else
				{
					rcv_len=0;
					state=CHK_STX;
					return 0;
				}
				break;
			case CHK_DATA:
			case CHK_ETX:
				if(len == 102)  //success
				{
					memset(MRZ1,0x00,45);
					memset(MRZ2,0x00,45);
					strncpy(MRZ1,(const char *)&rcv_packet[4],44);
					strncpy(MRZ2,(const char *)&rcv_packet[48],44);
					g_MRZ1 = (CString)MRZ1;
					g_MRZ2 = (CString)MRZ2;

					if(rcv_packet[104] == 0x03) //ETX
						state=CHK_LRC;
					else
					{
						rcv_len=0;
						state=CHK_STX;
						return 0;
					}

				}
				else
				{
					if(rcv_packet[5] == 0x03) //ETX
						state=CHK_LRC;
					else
					{
						rcv_len=0;
						state=CHK_STX;
						return 0;
					}
					return -1;
				}
				break;
			case CHK_LRC:
				rcv_len = (len == 102)?  (rcv_len-106) : (rcv_len-7);
				if(rcv_len)
				{
					int i = (len == 102)?  106:7;

					for(int j=0;j<rcv_len;j++,i++)
					{
						rcv_packet[j]=rcv_packet[i];
					}
				}
				run_f=0;
				state=CHK_STX;
				break;
		}
	}

	return 1;
}

extern "C" __declspec(dllexport) int __stdcall Scan()
{
	unsigned char packet[6];

	if( g_bCommOpened )
	{

		// Scanner Code send
		memset(packet,0x00,6);
		packet[0]=0x02;  //STX
		packet[1]=0x00;  //size
		packet[2]=0x02;  //size
		packet[3]=0xA1;  //Command ID
		packet[4]=0x03;  //ETX
		for(int i=1;i<5;i++)
			packet[5]^=packet[i];

		gCom.SendData( packet,6);

		return 1;
	}

	return 0;
}

extern "C" __declspec(dllexport) int __stdcall ScanCancel()
{
	int nRet = 0;
	unsigned char packet[6];

	if( g_bCommOpened )
	{
		// Scanner Code send
		memset(packet,0x00,6);
		packet[0]=0x02;  //STX
		packet[1]=0x00;  //size
		packet[2]=0x02;  //size
		packet[3]=0xAC;  //Command ID
		packet[4]=0x03;  //ETX
		for(int i=1;i<5;i++)
			packet[5]^=packet[i];

		gCom.SendData( packet,6);

		ExceptionReceiveData();

		return 1;
	}

	return 0;
}

extern "C" __declspec(dllexport) int __stdcall ReceiveData(int TimeOut)
{
	int nRetry = 0;
	int nRet = 0;
	int nReturn = 0;

	while(1){

		g_MRZ1.Empty();
		g_MRZ2.Empty();

		nRet = ReceiveCommData();
		if(nRet == 1){
			/////////////////////////////////////
			// 정상 데이터 수신
			/////////////////////////////////////
			nReturn = 1;
			break;
		}else if(nRet == -1){
			/////////////////////////////////////
			// 오류 데이터 수신
			/////////////////////////////////////
			Sleep(100);

			ExceptionReceiveData();

			nReturn = -1;
			return nReturn;
		}

		if(nRetry == (TimeOut*10)){

			ScanCancel();

			Sleep(100);

			ExceptionReceiveData();

			nReturn = 0;
			break;
		}

		Sleep(100);
		nRetry++;
	}

	return nReturn;
}


extern "C" __declspec(dllexport) int __stdcall GetMRZ1(char *refMRZ1)
{
	char MRZ1[44+1];

	memset(MRZ1,	0x00,	sizeof(MRZ1));
	memcpy(MRZ1, (LPSTR)(LPCTSTR)g_MRZ1,		g_MRZ1.GetLength());

	memcpy(refMRZ1,	MRZ1,	g_MRZ1.GetLength());

	return 1;

}

extern "C" __declspec(dllexport) int __stdcall GetMRZ2(char *refMRZ2)
{
	char MRZ2[44+1];

	memset(MRZ2,	0x00,	sizeof(MRZ2));
	memcpy(MRZ2, (LPSTR)(LPCTSTR)g_MRZ2,		g_MRZ2.GetLength());

	memcpy(refMRZ2,	MRZ2,	g_MRZ2.GetLength());

	return 1;
}

extern "C" __declspec(dllexport) int __stdcall Clear()
{
	g_MRZ1.Empty();
	g_MRZ2.Empty();
	
	return 1;
}

extern "C" __declspec(dllexport) int __stdcall ClosePort()
{
	if( g_bCommOpened ){
		gCom.ClosePort();

		gCom.Initialize();
		g_bCommOpened = FALSE;

		::DeleteCriticalSection(&g_cx);
	}
	
	return 1;
}

extern "C" __declspec(dllexport) int __stdcall GetPassportInfo(char *refPassInfo)
{
	char PassportInfo[65+1];

	memset(PassportInfo,	0x00,	sizeof(PassportInfo));
	memset(PassportInfo,	0x20,	65);

	if(g_MRZ1.GetLength() !=0 && g_MRZ2.GetLength() != 0)
	{
		CString tmpName = g_MRZ1.Mid(5, g_MRZ1.GetLength());
		int tmpPos = tmpName.Find("<<", 0);
		CString firstName = tmpName.Mid(0, tmpPos) ;
		CString lastName = tmpName.Mid(tmpPos, tmpName.GetLength());
		lastName.Replace("<", "");

		CString fullName;
		fullName.Format(_T("%s %s"), (LPCTSTR)firstName, (LPCTSTR)lastName);

		CString passportNo = g_MRZ2.Mid(0, 9);
		passportNo.Replace("<", "");

		CString country = g_MRZ2.Mid(10, 3);
		country.Replace("<", "");

		CString dateOfbirth = g_MRZ2.Mid(13, 6);
		dateOfbirth.Replace("<", "");

		CString gender = g_MRZ2.Mid(20, 1);
		gender.Replace("<", "");

		CString expiredate = g_MRZ2.Mid(21, 6);
		expiredate.Replace("<", "");

		memcpy(PassportInfo,	(LPCTSTR)fullName.GetBuffer(0), fullName.GetLength());
		memcpy(PassportInfo+40,	(LPCTSTR)passportNo.GetBuffer(0), passportNo.GetLength());
		memcpy(PassportInfo+49,	(LPCTSTR)country.GetBuffer(0), country.GetLength());
		memcpy(PassportInfo+52,	(LPCTSTR)gender.GetBuffer(0), gender.GetLength());
		memcpy(PassportInfo+53,	(LPCTSTR)dateOfbirth.GetBuffer(0), dateOfbirth.GetLength());
		memcpy(PassportInfo+59,	(LPCTSTR)expiredate.GetBuffer(0), expiredate.GetLength());

		memcpy(refPassInfo,	PassportInfo,	strlen(PassportInfo));

		return 1;
	}else{
		return 0;
	}
}

extern "C" __declspec(dllexport) int __stdcall CheckReceiveData()
{
	int nRet = 0;
	int nReturn = 0;

	nRet = ReceiveCommData();
	if(nRet == 1){
		/////////////////////////////////////
		// 정상 데이터 수신
		/////////////////////////////////////
		nReturn = 1;
	}else if(nRet == -1){
		/////////////////////////////////////
		// 오류 데이터 수신
		/////////////////////////////////////
		nReturn = -1;
	}else if(nRet == 0){
		/////////////////////////////////////
		// 데이터 미수신
		/////////////////////////////////////
		nReturn = 0;
	}

	return nReturn;
}
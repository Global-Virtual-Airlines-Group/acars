#include "fs9gauges.h"
#include "dsd_ipc2xml.h"
#include <MMsystem.h>
#include <math.h>
#include <iostream>
#include <string>
#include "FSUIPC_User.h"
using namespace std;

typedef struct _ACARS_Data
{
	BYTE Mode;			//0x7a20 1
	BYTE Connected;		//0x7a21 1
	DWORD Flight_ID;	//0x7a22 4
	char Pilot_ID[10];	//0x7a26 10	
	char Flight[10];	//0x7a30 10
	BYTE Leg;			//0x7a3a 1
	BYTE Chat;			//0x7a3b 1
	BYTE Update_COM;	//0x7a3c 1
	BYTE COM_Update;	//0x7a3d 1
	BYTE COM_Swap;		//0x7a3e 1
} ACARS;

_ACARS_Data Acars;

BYTE Mem[100];
DWORD Size = 100;
DWORD FSReq = SIM_ANY;
DWORD Result;

ID id_Mode, id_Connected, id_Flight, id_Leg, id_Chat, id_COM, id_Update, id_Swap;
ID id_Address;

int ipc_open = 0;
int i=0, flip=0;

#define  GAUGE_NAME  "interface\0"
#define   GAUGEHDR_VAR_NAME  gaugehdr_ipc2xml
#define   GAUGE_W    10
#include "dsd_ipc2xmlG.cpp"
#undef	GAUGE_NAME
#undef	GAUGEHDR_VAR_NAME
#undef	GAUGE_W

#define  GAUGE_NAME  "display\0"
#define   GAUGEHDR_VAR_NAME  gaugehdr_display
#define   GAUGE_W    70
#include "flight_displayG.cpp"

void FSAPI	module_init(void){}		

void FSAPI	module_deinit(void){}	

BOOL WINAPI DllMain (HINSTANCE hDLL, DWORD dwReason, LPVOID lpReserved)	
{														
	return TRUE;
}

GAUGESIMPORT	ImportTable =							
{														
	{ 0x0000000F, (PPANELS)NULL },
	{ 0x00000000, NULL }								
};

GAUGESLINKAGE	Linkage =								
{														
	0x00000013,											
	module_init,										
	module_deinit,										
	0,													
	0,
	FS9LINK_VERSION, 
	{
		(&gaugehdr_ipc2xml),
		(&gaugehdr_display),
		0 
	}
};


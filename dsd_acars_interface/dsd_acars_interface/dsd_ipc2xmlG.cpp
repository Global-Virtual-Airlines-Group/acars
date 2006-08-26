char ipc2xml_gauge_name[]  = GAUGE_NAME;
extern PELEMENT_HEADER   ipc2xml_list;
GAUGE_CALLBACK     gaugecall;
GAUGE_HEADER_FS700(GAUGE_W, ipc2xml_gauge_name, &ipc2xml_list, NULL, gaugecall, 0, 0, 0);

/*
L:ACARS_Mode
L:ACARS_Update_COM
*/

ID id_Mode_value, id_Mode_set, id_Update_COM_value, id_Update_COM_set;
BYTE Mode_value, Update_COM_value;
int Mode_set, Update_COM_set;
int process_please = 0;

void FSAPI gaugecall(PGAUGEHDR pgauge, int service_id, UINT32 extra_data)
{
switch(service_id)
{
case PANEL_SERVICE_CONNECT_TO_WINDOW:
	id_Mode = register_named_variable("ACARS_Mode");
	id_Connected = register_named_variable("ACARS_Connected");
	id_Flight = register_named_variable("ACARS_Flight_ID");
	id_Leg = register_named_variable("ACARS_Leg");
	id_Chat = register_named_variable("ACARS_Chat");
	id_COM = register_named_variable("ACARS_Update_COM");
	id_Update = register_named_variable("ACARS_COM_Radio_Update");
	id_Swap = register_named_variable("ACARS_COM_Radio_Swap");
	id_Mode_value = register_named_variable("ACARS_Mode_value");
	id_Mode_set = register_named_variable("ACARS_Mode_set");
	id_Update_COM_value = register_named_variable("ACARS_Update_COM_value");
	id_Update_COM_set = register_named_variable("ACARS_Update_COM_set");
break;

case PANEL_SERVICE_PRE_INITIALIZE:
	id_Mode = check_named_variable("ACARS_Mode");
	id_Connected = check_named_variable("ACARS_Connected");
	id_Flight = check_named_variable("ACARS_Flight_ID");
	id_Leg = check_named_variable("ACARS_Leg");
	id_Chat = check_named_variable("ACARS_Chat");
	id_COM = check_named_variable("ACARS_Update_COM");
	id_Update = check_named_variable("ACARS_COM_Radio_Update");
	id_Swap = check_named_variable("ACARS_COM_Radio_Swap");
	id_Mode_value = check_named_variable("ACARS_Mode_value");
	id_Mode_set = check_named_variable("ACARS_Mode_set");
	id_Update_COM_value = check_named_variable("ACARS_Update_COM_value");
	id_Update_COM_set = check_named_variable("ACARS_Update_COM_set");
break;

case PANEL_SERVICE_PRE_INSTALL:
	if ( ipc_open == 0 )
		{
		FSUIPC_Open2(FSReq, &Result, Mem, Size);
		FSUIPC_Write(0x8001, sizeof("X42WE5H2MT2Adva_acars_interface.gau"), "X42WE5H2MT2Adva_acars_interface.gau", &Result);
		FSUIPC_Process(&Result);
		if (Result == 0 ) ipc_open = 1;
		}
break;

case PANEL_SERVICE_PRE_UPDATE:

	if ( ipc_open == 0 )
		{
		FSUIPC_Open2(FSReq, &Result, Mem, Size);
		FSUIPC_Write(0x8001, sizeof("X42WE5H2MT2Adva_acars_interface.gau"), "X42WE5H2MT2Adva_acars_interface.gau", &Result);
		FSUIPC_Process(&Result);
		if (Result == 0 ) ipc_open = 1;
		}

	if ( ipc_open == 1 && flip >= 3 )
		{
		FSUIPC_Read(0x7A20, sizeof(_ACARS_Data), &Acars.Mode, &Result);
		FSUIPC_Process(&Result);
		flip = 0;
		set_named_variable_value(id_Mode, (float)Acars.Mode);
		set_named_variable_value(id_Connected, (float)Acars.Connected);
		set_named_variable_value(id_Flight, (float)Acars.Flight_ID);
		set_named_variable_value(id_Leg, (float)Acars.Leg);
		set_named_variable_value(id_Chat, (float)Acars.Chat);
		set_named_variable_value(id_COM, (float)Acars.Update_COM);
		set_named_variable_value(id_Update, (float)Acars.COM_Update);
		set_named_variable_value(id_Swap, (float)Acars.COM_Swap);
		}
	else flip++;

	Mode_set = (int)get_named_variable_value(id_Mode_set);
	if (Mode_set)
		{
		Mode_value = (BYTE)get_named_variable_value(id_Mode_value);
		FSUIPC_Write(0x7a20, 1, &Mode_value, &Result);
		process_please = 1;
		set_named_variable_value(id_Mode_set, 0.);
		}
	Update_COM_set = (int)get_named_variable_value(id_Update_COM_set);
	if (Update_COM_set)
		{
		Update_COM_value = (BYTE)get_named_variable_value(id_Update_COM_value);
		FSUIPC_Write(0x7a3c, 1, &Update_COM_value, &Result);
		process_please = 1;
		set_named_variable_value(id_Update_COM_set, 0.);
		}
	if (process_please)
		{
		FSUIPC_Process(&Result);
		process_please = 0;
		}

break;

case PANEL_SERVICE_DISCONNECT:
	FSUIPC_Close();
break;

}
}

MAKE_STATIC(Static1,Rec0,NULL,NULL,IMAGE_USE_ERASE | IMAGE_USE_TRANSPARENCY | BIT7,0,0,0)
PELEMENT_HEADER ipc2xml_list = &Static1.header;

char Flight_Display_gauge_name[] = GAUGE_NAME;
extern PELEMENT_HEADER   Flight_Display_list;
GAUGE_HEADER_FS700(GAUGE_W, Flight_Display_gauge_name, &Flight_Display_list, NULL, NULL, 0, 0, 0);

#define GAUGE_CHARSET2           DEFAULT_CHARSET
#define GAUGE_FONT_DEFAULT2          "Arial"
#define GAUGE_WEIGHT_DEFAULT2     FW_BOLD

FLOAT64 FSAPI callback3( PELEMENT_STRING pelement)
{
sprintf(pelement->string,"%s",Acars.Flight);
return 0;
}

FLOAT64 FSAPI callback2( PELEMENT_STRING pelement)
{
sprintf(pelement->string,"%s",Acars.Pilot_ID);
return 0;
}

MAKE_STRING(String3,NULL,NULL,IMAGE_USE_ERASE | IMAGE_USE_TRANSPARENCY | BIT7,0,60,20,85,17,20,MODULE_VAR_NONE,MODULE_VAR_NONE,MODULE_VAR_NONE,RGB(255,255,255),RGB(0,0,0),RGB(0,0,0),GAUGE_FONT_DEFAULT2,GAUGE_WEIGHT_DEFAULT2,GAUGE_CHARSET2,0,0,NULL,callback3)
PELEMENT_HEADER ElementList2[] = {&String3.header,NULL};

MAKE_STRING(String2,&ElementList2,NULL,IMAGE_USE_ERASE | IMAGE_USE_TRANSPARENCY | BIT7,0,60,2,85,17,20,MODULE_VAR_NONE,MODULE_VAR_NONE,MODULE_VAR_NONE,RGB(255,255,255),RGB(0,0,0),RGB(0,0,0),GAUGE_FONT_DEFAULT2,GAUGE_WEIGHT_DEFAULT2,GAUGE_CHARSET2,0,0,NULL,callback2)
PELEMENT_HEADER ElementList3[] = {&String2.header,NULL};

MAKE_STATIC(Display_Static,Rec1,&ElementList3,NULL,IMAGE_USE_ERASE | IMAGE_USE_TRANSPARENCY | BIT7,0,0,0)
PELEMENT_HEADER Flight_Display_list = &Display_Static.header;

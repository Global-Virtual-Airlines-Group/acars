VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DVA ACARS"
   ClientHeight    =   8715
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Timer tmrStartCheck 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8640
      Top             =   3240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Tahoma"
   End
   Begin VB.Timer tmrFlightTime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8160
      Top             =   4200
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   8640
      Top             =   3720
   End
   Begin VB.Timer tmrPosUpdates 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   8160
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock wsckMain 
      Left            =   8640
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnectDisconnect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8140
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   215
      Width           =   1235
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   29
      Top             =   8415
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5821
            MinWidth        =   5821
            Text            =   "Status: Connected to ACARS server"
            TextSave        =   "Status: Connected to ACARS server"
            Key             =   "status"
            Object.Tag             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Phase: Airborne"
            TextSave        =   "Phase: Airborne"
            Key             =   "flightphase"
            Object.Tag             =   "flightphase"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4340
            MinWidth        =   4340
            Text            =   "Last Position Report: 02:38:08 Z"
            TextSave        =   "Last Position Report: 02:38:08 Z"
            Key             =   "lastposrep"
            Object.Tag             =   "lastposrep"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5644
            MinWidth        =   5644
            Text            =   "Flight Time: 00:00:00"
            TextSave        =   "Flight Time: 00:00:00"
            Key             =   "simtime"
            Object.Tag             =   "simtime"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdPIREP 
      Caption         =   "File PIREP"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8140
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   970
      Visible         =   0   'False
      Width           =   1235
   End
   Begin VB.CommandButton cmdStartStopFlight 
      Caption         =   "Start Flight"
      Height          =   345
      Left            =   8140
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   595
      Width           =   1235
   End
   Begin TabDlg.SSTab SSTab1 
      CausesValidation=   0   'False
      Height          =   3720
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   6562
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "ACARS Messages"
      TabPicture(0)   =   "frmMain.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "rtfText"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSend"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCmd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Connected Pilots"
      TabPicture(1)   =   "frmMain.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBusy"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdUpdatePilotList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "infoFrame"
      Tab(1).Control(3)=   "lstPilots"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Air Traffic Control"
      TabPicture(2)   =   "frmMain.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lstATC"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "ctrFrame"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "radioFrame"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "XML Message Data"
      TabPicture(3)   =   "frmMain.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "rtfDebug"
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdBusy 
         Caption         =   "I'm Busy"
         Height          =   255
         Left            =   -70920
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1815
      End
      Begin RichTextLib.RichTextBox rtfDebug 
         Height          =   3135
         Left            =   -74880
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   120
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5530
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmMain.frx":007C
      End
      Begin VB.Frame radioFrame 
         Caption         =   "Communication Radio Frequencies"
         Height          =   1575
         Left            =   4440
         TabIndex        =   46
         Top             =   1695
         Width           =   4620
         Begin VB.CommandButton cmdSetAPP 
            Caption         =   "Set to xxx.xx"
            Height          =   290
            Left            =   3070
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   1230
            Width           =   1350
         End
         Begin VB.CommandButton cmdSetCTR 
            Caption         =   "Set to xxx.xx"
            Height          =   290
            Left            =   3070
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   910
            Width           =   1350
         End
         Begin VB.CommandButton cmdSetDEP 
            Caption         =   "Set to xxx.xx"
            Height          =   290
            Left            =   3070
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   590
            Width           =   1350
         End
         Begin VB.CommandButton cmdSetGND 
            Caption         =   "Set to xxx.xx"
            Height          =   290
            Left            =   3070
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   270
            Width           =   1350
         End
         Begin VB.CommandButton cmdTuneAPP 
            Caption         =   "Set COM1"
            Height          =   290
            Left            =   2040
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   1230
            Width           =   900
         End
         Begin VB.CommandButton cmdTuneCTR 
            Caption         =   "Set COM1"
            Height          =   290
            Left            =   2040
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   910
            Width           =   900
         End
         Begin VB.CommandButton cmdTuneDEP 
            Caption         =   "Set COM1"
            Height          =   290
            Left            =   2040
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   590
            Width           =   900
         End
         Begin VB.CommandButton cmdTuneGND 
            Caption         =   "Set COM1"
            Height          =   290
            Left            =   2040
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   270
            Width           =   900
         End
         Begin VB.TextBox txtAPPfreq 
            Height          =   290
            Left            =   1200
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1220
            Width           =   700
         End
         Begin VB.TextBox txtCTRfreq 
            Height          =   290
            Left            =   1200
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   900
            Width           =   700
         End
         Begin VB.TextBox txtDEPfreq 
            Height          =   290
            Left            =   1200
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   590
            Width           =   700
         End
         Begin VB.TextBox txtGNDfreq 
            Height          =   290
            Left            =   1200
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   270
            Width           =   700
         End
         Begin VB.Label radoAPPlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Approach"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   1240
            Width           =   1035
         End
         Begin VB.Label radioCTRlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Center"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   940
            Width           =   1030
         End
         Begin VB.Label radioDEPlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Departure"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   620
            Width           =   1030
         End
         Begin VB.Label radioGNDlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Ground/Tower"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   330
            Width           =   1030
         End
      End
      Begin VB.Frame ctrFrame 
         Caption         =   "Controller Information"
         Height          =   1545
         Left            =   4440
         TabIndex        =   45
         Top             =   90
         Visible         =   0   'False
         Width           =   4620
         Begin VB.CommandButton cmdUpdateATCList 
            Caption         =   "Update Air Traffic Control List"
            Height          =   270
            Left            =   840
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1200
            Width           =   3015
         End
         Begin VB.Label atcInfoFreq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "xxx.xx"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3480
            TabIndex        =   66
            Top             =   900
            Width           =   735
         End
         Begin VB.Label atcInfoRating 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Senior Controller"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   65
            Top             =   900
            Width           =   1335
         End
         Begin VB.Label atcInfoFacility 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "KZTL_V_CTR (Center)"
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   1080
            TabIndex        =   64
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label atcInfoName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Controller Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   63
            Top             =   285
            Width           =   3135
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "Frequency"
            Height          =   255
            Left            =   2520
            TabIndex        =   54
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "ATC Rating"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "Facility Info"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   615
            Width           =   855
         End
         Begin VB.Label label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Left            =   360
            TabIndex        =   51
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdUpdatePilotList 
         Caption         =   "Update Connected Pilot List"
         Height          =   270
         Left            =   -69000
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3015
      End
      Begin VB.ListBox lstATC 
         ForeColor       =   &H00800000&
         Height          =   3180
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   33
         Top             =   120
         Width           =   4215
      End
      Begin VB.Frame infoFrame 
         Caption         =   "Pilot Information"
         Height          =   2595
         Left            =   -70920
         TabIndex        =   32
         Top             =   240
         Width           =   4935
         Begin VB.CommandButton cmdBan 
            Caption         =   "Disconnect User and Block Address"
            Height          =   255
            Left            =   1920
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   2235
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CommandButton cmdKick 
            Caption         =   "Disconnect User"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   2235
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label pilotInfoRoute 
            BackStyle       =   0  'Transparent
            Caption         =   "Airport (XXXX) - Airport (XXXX)"
            Height          =   495
            Left            =   1080
            TabIndex        =   91
            Top             =   1395
            Width           =   3615
         End
         Begin VB.Label pilotInfoEqType 
            BackStyle       =   0  'Transparent
            Caption         =   "BXXX-XXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   255
            Left            =   2400
            TabIndex        =   90
            Top             =   1110
            Width           =   1335
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "Flight Route"
            Height          =   255
            Left            =   30
            TabIndex        =   89
            Top             =   1390
            Width           =   945
         End
         Begin VB.Label lblBusy 
            Caption         =   "BUSY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3720
            TabIndex        =   87
            Top             =   270
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label pilotInfoConnectionInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Build xxx from xxx.xxx.xxx.xxx"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   44
            Top             =   1920
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.Label pilotInfoConnectionLabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Connected"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label pilotInfoFlightCode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "DVAXXXX"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1080
            TabIndex        =   42
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Flying"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1110
            Width           =   855
         End
         Begin VB.Label pilotInfoFlightData 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "0 legs, 0.0 hours"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   40
            Top             =   810
            Width           =   1935
         End
         Begin VB.Label pilotInfoRank 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Rank, Equipment Type"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   39
            Top             =   540
            Width           =   2775
         End
         Begin VB.Label pilotInfoName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000016&
            BackStyle       =   0  'Transparent
            Caption         =   "Pilot Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1080
            TabIndex        =   38
            Top             =   250
            Width           =   2415
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "Flights"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   810
            Width           =   855
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Rank"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   540
            Width           =   735
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   260
            Width           =   855
         End
      End
      Begin VB.ListBox lstPilots 
         ForeColor       =   &H00800000&
         Height          =   3180
         Left            =   -74880
         Sorted          =   -1  'True
         TabIndex        =   31
         Top             =   120
         Width           =   3855
      End
      Begin VB.TextBox txtCmd 
         Height          =   285
         Left            =   -74880
         TabIndex        =   19
         Top             =   3000
         Width           =   7965
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "SEND"
         Height          =   285
         Left            =   -66840
         TabIndex        =   21
         Top             =   3000
         Width           =   930
      End
      Begin RichTextLib.RichTextBox rtfText 
         CausesValidation=   0   'False
         Height          =   2800
         Left            =   -74880
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   4948
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmMain.frx":00F7
      End
   End
   Begin VB.Frame authFrame 
      Caption         =   "Pilot Authentication"
      Height          =   720
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtPilotName 
         BackColor       =   &H8000000F&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   265
         Width           =   1575
      End
      Begin VB.TextBox txtPilotID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3915
         TabIndex        =   1
         Top             =   265
         WhatsThisHelpID =   104
         Width           =   950
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   6435
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   265
         Width           =   1090
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pilot Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   295
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Pilot ID:"
         Height          =   255
         Left            =   3195
         TabIndex        =   23
         Top             =   295
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   5460
         TabIndex        =   22
         Top             =   295
         Width           =   915
      End
   End
   Begin VB.Frame flightInfoFrame 
      Caption         =   "Flight Information"
      Height          =   3615
      Left            =   120
      TabIndex        =   75
      Top             =   960
      Width           =   7935
      Begin VB.ComboBox cboAirline 
         Height          =   315
         ItemData        =   "frmMain.frx":0172
         Left            =   1395
         List            =   "frmMain.frx":0179
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtAirportL 
         Height          =   315
         Left            =   4800
         TabIndex        =   14
         Top             =   1320
         Width           =   530
      End
      Begin VB.TextBox txtAirportA 
         Height          =   315
         Left            =   4800
         TabIndex        =   10
         Top             =   960
         Width           =   530
      End
      Begin VB.TextBox txtAirportD 
         Height          =   315
         Left            =   4800
         TabIndex        =   8
         Top             =   600
         Width           =   530
      End
      Begin VB.TextBox txtRoute 
         Height          =   675
         Left            =   1395
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1740
         Width           =   6345
      End
      Begin VB.TextBox txtRemarks 
         Height          =   615
         Left            =   1395
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   2520
         Width           =   6345
      End
      Begin VB.TextBox txtFlightNumber 
         Height          =   300
         Left            =   4120
         TabIndex        =   4
         Top             =   240
         Width           =   500
      End
      Begin VB.TextBox txtCruiseAlt 
         Height          =   300
         Left            =   6675
         TabIndex        =   11
         Top             =   600
         Width           =   1080
      End
      Begin VB.ComboBox cboEquipment 
         Height          =   315
         ItemData        =   "frmMain.frx":0189
         Left            =   6455
         List            =   "frmMain.frx":018B
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1320
      End
      Begin VB.ComboBox cboAirportD 
         CausesValidation=   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":018D
         Left            =   1395
         List            =   "frmMain.frx":0194
         TabIndex        =   7
         Text            =   "cboAirportD"
         ToolTipText     =   "This is the Airport you are departing from."
         Top             =   600
         Width           =   3335
      End
      Begin VB.ComboBox cboAirportA 
         Height          =   315
         ItemData        =   "frmMain.frx":01A4
         Left            =   1395
         List            =   "frmMain.frx":01AB
         TabIndex        =   9
         Text            =   "cboAirportA"
         ToolTipText     =   "This is the Airport you are arriving at."
         Top             =   960
         Width           =   3335
      End
      Begin VB.ComboBox cboNetwork 
         Height          =   315
         ItemData        =   "frmMain.frx":01BB
         Left            =   6675
         List            =   "frmMain.frx":01CB
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   1090
      End
      Begin VB.ComboBox cboAirportL 
         Height          =   315
         ItemData        =   "frmMain.frx":01EB
         Left            =   1395
         List            =   "frmMain.frx":01ED
         TabIndex        =   13
         Text            =   "cboAirportL"
         Top             =   1320
         Width           =   3335
      End
      Begin VB.CheckBox chkCheckRide 
         Caption         =   "Aircraft Check Ride"
         Height          =   255
         Left            =   5760
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1360
         Width           =   1800
      End
      Begin VB.TextBox txtLeg 
         Height          =   300
         Left            =   5040
         TabIndex        =   5
         Text            =   "1"
         Top             =   240
         Width           =   280
      End
      Begin MSComctlLib.ProgressBar PositionProgress 
         Height          =   255
         Left            =   3900
         TabIndex        =   76
         Top             =   3240
         Visible         =   0   'False
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Airline:"
         Height          =   255
         Left            =   595
         TabIndex        =   92
         Top             =   260
         Width           =   695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Flight"
         Height          =   255
         Left            =   3560
         TabIndex        =   86
         Top             =   285
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Flight Route:"
         Height          =   255
         Left            =   100
         TabIndex        =   85
         Top             =   1740
         Width           =   1190
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks:"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   2475
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Equipment:"
         Height          =   255
         Left            =   5510
         TabIndex        =   83
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Cruise Altitude:"
         Height          =   255
         Left            =   5400
         TabIndex        =   82
         Top             =   645
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Arriving at:"
         Height          =   255
         Left            =   195
         TabIndex        =   81
         Top             =   1005
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Network:"
         Height          =   255
         Left            =   5880
         TabIndex        =   80
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label progressLabel 
         Caption         =   "FILING OFFLINE FLIGHT REPORT"
         Height          =   255
         Left            =   1410
         TabIndex        =   79
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Alternate:"
         Height          =   255
         Left            =   195
         TabIndex        =   78
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Leg"
         Height          =   255
         Left            =   4640
         TabIndex        =   20
         Top             =   285
         Width           =   345
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Departing from:"
         Height          =   255
         Left            =   75
         TabIndex        =   77
         Top             =   645
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFlight 
      Caption         =   "&Flight"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuFlightSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFlightPlan 
         Caption         =   "&Open Flight Plan"
      End
      Begin VB.Menu mnuSaveFlightPlan 
         Caption         =   "&Save SB3 Flight Plan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFlightSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearChatText 
         Caption         =   "Clear Chat Text"
      End
      Begin VB.Menu mnuFlightSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlightExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCOM 
      Caption         =   "&Communications"
      Begin VB.Menu mnuCOMPlaySoundForChat 
         Caption         =   "Play &Sound For Chat"
      End
      Begin VB.Menu mnuCOMHideBusy 
         Caption         =   "Hide &Messages when Busy"
      End
      Begin VB.Menu mnuCOMSterileCockpit 
         Caption         =   "Sterile &Cockpit"
      End
      Begin VB.Menu mnuCOMFreezeWindow 
         Caption         =   "&Freeze Message Window"
      End
      Begin VB.Menu mnuCOMShowTimestamps 
         Caption         =   "Show &Timestamps"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsFlyOffline 
         Caption         =   "Fly &Disconnected"
      End
      Begin VB.Menu mnuOptionsSavePassword 
         Caption         =   "Save &Password"
      End
      Begin VB.Menu mnuOptionsShowPilotNames 
         Caption         =   "Show Pilot &Names"
      End
      Begin VB.Menu mnuOptionsHide 
         Caption         =   "&Hide when Minimized"
      End
      Begin VB.Menu mnuOptionsShowDebugMessages 
         Caption         =   "Show Debug Messages"
      End
      Begin VB.Menu mnuOptionsDisableAutoSave 
         Caption         =   "Disable &Automatic Saves"
      End
      Begin VB.Menu mnuOptionsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsWideFS 
         Caption         =   "&WideFS Installed"
      End
      Begin VB.Menu mnuOptionsSB3Support 
         Caption         =   "&SquawkBox 3 Integration"
      End
      Begin VB.Menu mnuOptionsTS2Support 
         Caption         =   "&TeamSpeak 2 Integration"
      End
      Begin VB.Menu mnuOptionsGaugeIntegration 
         Caption         =   "&Gauge Integration"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About ACARS"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "SystemTray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSystrayRestore 
         Caption         =   "Display ACARS"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CmdHistory() As String
Dim intCmdHistIndex As Integer
Dim blnFirstCmd As Boolean
Dim strCmdBuffer As String
Dim intCmdBufferSel As Integer

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button upt

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

Private Sub cboAirportA_Click()
    Set info.AirportA = config.GetAirportByOfs(cboAirportA.ListIndex - 1)
    If Not (info.AirportA Is Nothing) Then txtAirportA.Text = info.AirportA.ICAO
End Sub

Private Sub cboAirportD_Click()
    Set info.airportD = config.GetAirportByOfs(cboAirportD.ListIndex - 1)
    If Not (info.airportD Is Nothing) Then txtAirportD.Text = info.airportD.ICAO
End Sub

Private Sub cboAirportL_Click()
    Set info.AirportL = config.GetAirportByOfs(cboAirportL.ListIndex - 1)
    If Not (info.AirportL Is Nothing) Then txtAirportL.Text = info.AirportL.ICAO
End Sub

Private Sub cboAirline_Click()
    Set info.Airline = config.GetAirlineByOfs(cboAirline.ListIndex - 1)
End Sub

Private Sub cboEquipment_Click()
    info.EquipmentType = cboEquipment.List(cboEquipment.ListIndex)
End Sub

Private Sub cboNetwork_Click()
    info.Network = cboNetwork.List(cboNetwork.ListIndex)
End Sub

Private Sub chkCheckRide_Click()
    info.CheckRide = Not (chkCheckRide.value = 0)
End Sub

Private Sub cmdBan_Click()
    Dim ID As String
    Dim p As Pilot
    
    'Check access
    If Not config.HasRole("HR") Then Exit Sub
    
    'Get the pilot
    ID = Split(lstPilots.List(lstPilots.ListIndex), " ")(0)
    Set p = users.GetPilot(ID)
    If (p Is Nothing) Then Exit Sub
    
    'Send the kick request
    SendKick p.RemoteAddress, True
    ReqStack.Send
End Sub

Private Sub cmdBusy_Click()
    If config.ACARSConnected Then
        config.Busy = Not config.Busy
        SendBusy config.Busy
        ReqStack.Send
        
        If config.Busy Then
            cmdBusy.Caption = "Available for Chat"
        Else
            cmdBusy.Caption = "I'm Busy"
        End If
    End If
End Sub

Private Sub cmdConnectDisconnect_Click()
    ToggleACARSConnection
End Sub

Private Sub cmdKick_Click()
    Dim ID As String
    Dim p As Pilot
    
    'Check access
    If Not config.HasRole("HR") Then Exit Sub
    
    'Get the pilot
    ID = Split(lstPilots.List(lstPilots.ListIndex), " ")(0)
    Set p = users.GetPilot(ID)
    If (p Is Nothing) Then Exit Sub
    
    'Send the kick request
    SendKick p.ID, False
    ReqStack.Send
End Sub

Private Sub cmdPIREP_Click()
    Dim msgID As Long
    
    'Set critical error handler
    On Error GoTo FatalError

    'If we're not connected, then stop
    If Not config.ACARSConnected Then
        MsgBox "You must be connected to the ACARS server to file a Flight Report.", vbExclamation + vbOKOnly, "Not Connected"
        Exit Sub
    End If
    
    'Make sure that there is data
    If Not info.FlightData Then
        MsgBox "The Flight must be completed before you can file a Flight Report.", vbExclamation + vbOKOnly, "Incomplete Flight"
        Exit Sub
    End If
    
    'If we've already filed the PIREP and it was approved, then stop
    If info.PIREPFiled Then Exit Sub
    
    'Disable the option
    frmMain.cmdPIREP.enabled = False
    
    'If we don't have a flight ID, then file the flight info
    If (info.FlightID = 0) Then
        info.InfoReqID = SendFlightInfo(info)
        ReqStack.Send
        ShowMessage "Sending Offline Flight Information", ACARSTEXTCOLOR
    
        'If we time out, raise an error
        If Not WaitForACK(msgID, 7500) Then
            MsgBox "ACARS Server timed out returning Flight ID!", vbCritical + vbOKOnly
            frmMain.cmdPIREP.enabled = True
            Exit Sub
        End If
    End If
    
    'If we have offline position reports, file them
    If Positions.HasData Then
        Dim Queue As Variant
        Dim x As Integer
    
        'Display the progress bar
        Queue = Positions.Queue
        frmMain.progressLabel.visible = True
        frmMain.PositionProgress.value = 0
        frmMain.PositionProgress.Max = UBound(Queue) + 2
        frmMain.PositionProgress.visible = True
        
        'File the Flight Positions
        ShowMessage "Sending " + CStr(UBound(Queue) + 1) + " Position Records", ACARSTEXTCOLOR
        For x = 0 To UBound(Queue)
            msgID = SendPosition(Queue(x), True)
            frmMain.PositionProgress.value = x
            If ((x Mod 4) = 0) Then
                ReqStack.Send
                frmMain.PositionProgress.Refresh
                Call WaitForACK(msgID, 3000)
            End If
        Next

        'Clear the offline queue
        Call Positions.Clear
    End If
    
    'End the flight (the ACARS server will discard multiple messages)
    msgID = SendEndFlight
    ReqStack.Send
    Call WaitForACK(msgID, 1000)

    'Send the PIREP
    msgID = SendPIREP(info)
    frmMain.PositionProgress.value = frmMain.PositionProgress.Max
    ShowMessage "Sending Flight Report " + Hex(msgID), ACARSTEXTCOLOR
    ReqStack.Send
    
    'Wait for the ACK
    If Not WaitForACK(msgID, 15000) Then
        MsgBox "ACARS Server timed out sending Flight Report!", vbCritical + vbOKOnly, "Time Out"
        cmdPIREP.enabled = True
        progressLabel.visible = False
        Exit Sub
    End If
    
    'Kill the persisted data
    DeleteSavedFlight SavedFlightID(info)
    config.SaveFlightCode ""
    
    'Reset the flight data and enable the fields
    txtFlightNumber.Text = ""
    txtLeg.Text = "1"
    txtCruiseAlt.Text = ""
    txtRoute.Text = ""
    txtRemarks.Text = ""
    cboAirportD.ListIndex = 0
    cboAirportA.ListIndex = 0
    cboAirportL.ListIndex = 0
    cboNetwork.ListIndex = 0
    cboEquipment.ListIndex = 0
    chkCheckRide.value = 0
    LockFlightInfo True
    
    'Update buttons/info
    Set info = New FlightData
    config.UpdateFlightInfo
    cmdStartStopFlight.enabled = True
    cmdPIREP.visible = False
    cmdPIREP.enabled = False
    progressLabel.visible = False
    PositionProgress.visible = False
    
    'Show status messages
    ShowFSMessage "Flight Report Successfully filed", True, 15
    MsgBox "Flight Report filed Successfully.", vbInformation + vbOKOnly, "Flight Report Filed"
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Error filing Flight Report " + err.Description + " at Line " + CStr(Erl), vbCritical + _
        vbOKOnly, "Flight Report Error"
    Resume ExitSub
    
End Sub

Private Sub cmdSend_Click()
    txtCmd_KeyPress 13
End Sub

Private Sub cmdSetAPP_Click()
    txtAPPfreq.Text = atcInfoFreq.Caption
End Sub

Private Sub cmdSetCTR_Click()
    txtCTRfreq.Text = atcInfoFreq.Caption
End Sub

Private Sub cmdSetDEP_Click()
    txtDEPfreq.Text = atcInfoFreq.Caption
End Sub

Private Sub cmdSetGND_Click()
    txtGNDfreq.Text = atcInfoFreq.Caption
End Sub

Private Sub cmdStartStopFlight_Click()
    If info.InFlight Then
        StopFlight (Not config.FSUIPCConnected)
    ElseIf (info.FlightPhase = ERROR) Then
        RestoreFlight
    Else
        StartFlight
    End If
End Sub

Private Sub cmdTuneAPP_Click()
    SetCOM1 txtAPPfreq.Text
End Sub

Private Sub cmdTuneCTR_Click()
    SetCOM1 txtCTRfreq.Text
End Sub

Private Sub cmdTuneDEP_Click()
    SetCOM1 txtDEPfreq.Text
End Sub

Private Sub cmdTuneGND_Click()
    SetCOM1 txtGNDfreq.Text
End Sub

Private Sub cmdUpdateATCList_Click()
    If config.ACARSConnected Then
        RequestATCInfo LCase(info.Network)
        ReqStack.Send
    End If
End Sub

Private Sub cmdUpdatePilotList_Click()
    If config.ACARSConnected Then
        RequestPilotList
        ReqStack.Send
    End If
End Sub

Private Sub Form_GotFocus()
    If config.MsgReceived And config.FSUIPCConnected Then GAUGE_ClearChat
End Sub

Private Sub Form_Load()
    intCmdHistIndex = 0
    blnFirstCmd = True
    
    'Ensure timers are enabled
    tmrPing.enabled = False
    tmrPosUpdates.enabled = False
    tmrFlightTime.enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Msg As Long
    
    'Do nothing if not minimized
    If (WindowState <> vbMinimized) Then Exit Sub
    
    Msg = x / Screen.TwipsPerPixelX
    If (Msg = WM_LBUTTONDBLCLK) Then
        Call mnuSystrayRestore_Click
    ElseIf ((Msg = WM_RBUTTONUP) Or (Msg = WM_LBUTTONUP)) Then
        SetForegroundWindow Me.hWnd
        PopupMenu mnuTray
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ConfirmExit Then Cancel = 1
End Sub

Private Sub Form_Resize()
    If (config.HideWhenMinimized And (WindowState = vbMinimized)) Then
        AddIcon App.CompanyName + " ACARS", frmSplash.Icon
        TaskBarHide frmMain.hWnd
    ElseIf (config.HideWhenMinimized And (WindowState = vbNormal)) Then
        RemoveIcon
        TaskBarShow frmMain.hWnd
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (WindowState = vbMinimized) Then RemoveIcon

    'Close FSUIPC connection
    FSUIPC_Close
    config.FSUIPCConnected = False
    
    'Save settings
    config.Save

    'Unload all forms.
    Dim i As Integer
    For i = (Forms.Count - 1) To 0 Step -1
        Unload Forms(i)
    Next
    
    End
End Sub

Private Sub lstATC_Click()
    Dim ID As String
    Dim ctr As Controller
    
    'Get the controller
    ID = Split(lstATC.List(lstATC.ListIndex), " ")(0)
    Set ctr = users.GetATC(ID)
    If (ctr Is Nothing) Then Exit Sub
    
    'Update the controller info
    users.SelectedATC = ctr.ID
    atcInfoName.Caption = ctr.name
    atcInfoFacility.Caption = ctr.FacilityInfo
    atcInfoRating.Caption = ctr.Rating
    atcInfoFreq.Caption = ctr.Frequency
    ctrFrame.visible = True
    
    'Update buttons
    UpdateTuneButtons True, (ctr.Frequency <> "199.998"), "Set to " + ctr.Frequency
End Sub

Private Sub lstPilots_Click()
    Dim ID As String
    Dim p As Pilot
    
    'Get the pilot
    ID = Split(lstPilots.List(lstPilots.ListIndex), " ")(0)
    Set p = users.GetPilot(ID)
    If (p Is Nothing) Then Exit Sub
    
    'Update the pilot info
    users.SelectedPilot = p.ID
    pilotInfoName.Caption = p.name
    pilotInfoRank.Caption = p.Rank + ", " + p.EquipmentType
    pilotInfoFlightData.Caption = p.FlightTotals
    
    'Show flight data
    If (p.flightCode <> "") Then
        Label17.visible = True
        pilotInfoFlightCode.visible = True
        pilotInfoFlightCode.Caption = p.flightCode
        pilotInfoEqType.visible = True
        pilotInfoEqType.Caption = p.FlightEQ
        If Not ((p.AirportA Is Nothing) Or (p.airportD Is Nothing)) Then
            Label22.visible = True
            pilotInfoRoute.visible = True
            pilotInfoRoute.Caption = p.airportD.name + " (" + p.airportD.ICAO + ") - " + _
                p.AirportA.name + " (" + p.AirportA.ICAO + ")"
        Else
            Label22.visible = False
            pilotInfoRoute.visible = False
        End If
    Else
        Label17.visible = False
        pilotInfoFlightCode.visible = False
        pilotInfoEqType.visible = False
        Label22.visible = False
        pilotInfoRoute.visible = False
    End If
    
    lblBusy.visible = p.IsBusy
    infoFrame.visible = True
    If config.HasRole("HR") Then
        pilotInfoConnectionInfo.Caption = "Build " + CStr(p.ClientBuild) + " from " + p.RemoteAddress
        pilotInfoConnectionLabel.visible = True
        pilotInfoConnectionInfo.visible = True
        cmdKick.visible = True
        cmdBan.visible = True
    ElseIf pilotInfoConnectionLabel.visible Then
        pilotInfoConnectionLabel.visible = False
        pilotInfoConnectionInfo.visible = False
        cmdKick.visible = False
        cmdBan.visible = False
    End If
End Sub

Private Sub mnuClearChatText_Click()
    rtfText.Text = ""
End Sub

Private Sub mnuCOMFreezeWindow_Click()
    config.FreezeWindow = Not config.FreezeWindow
    config.UpdateSettingsMenu
End Sub

Private Sub mnuConnect_Click()
    ToggleACARSConnection
End Sub

Public Sub UpdateTuneButtons(visible As Boolean, enabled As Boolean, Optional btnCaption As String = "")
    Dim TuneButtons As Variant
    Dim btn As Variant
    
    TuneButtons = Array(cmdSetGND, cmdSetDEP, cmdSetCTR, cmdSetAPP)
    For Each btn In TuneButtons
        btn.Caption = btnCaption
        btn.visible = visible
        btn.enabled = enabled
    Next
End Sub

Private Sub mnuFlightExit_Click()
    Unload frmMain
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuOpenFlightPlan_Click()
    FPlan_Open
    config.UpdateFlightInfo
End Sub

Private Sub mnuOptionsDisableAutoSave_Click()
    config.DisableAutoSave = Not config.DisableAutoSave
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsGaugeIntegration_Click()
    config.GaugeSupport = Not config.GaugeSupport
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsHide_Click()
    config.HideWhenMinimized = Not config.HideWhenMinimized
    config.UpdateSettingsMenu
End Sub

Private Sub mnuCOMHideBusy_Click()
    config.HideMessagesWhenBusy = Not config.HideMessagesWhenBusy
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsSB3Support_Click()
    config.SB3Support = Not config.SB3Support
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsShowPilotNames_Click()
    config.ShowPilotNames = Not config.ShowPilotNames
    config.UpdateSettingsMenu
End Sub

Private Sub mnuCOMSterileCockpit_Click()
    config.SterileCockpit = Not config.SterileCockpit
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsTS2Support_Click()
    config.TS2Support = Not config.TS2Support
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsWideFS_Click()
    config.WideFSInstalled = Not config.WideFSInstalled
    config.UpdateSettingsMenu
End Sub

Private Sub mnuSaveFlightPlan_Click()
    SB3Plan_Save
End Sub

Private Sub mnuOptionsFlyOffline_Click()
    config.FlyOffline = Not config.FlyOffline
    config.UpdateSettingsMenu
End Sub

Private Sub mnuCOMPlaySoundForChat_Click()
    config.PlaySound = Not config.PlaySound
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsSavePassword_Click()
    config.SavePassword = Not config.SavePassword
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsShowDebugMessages_Click()
    config.ShowDebug = Not config.ShowDebug
    config.UpdateSettingsMenu
End Sub

Private Sub mnuCOMShowTimestamps_Click()
    config.ShowTimestamps = Not config.ShowTimestamps
    config.UpdateSettingsMenu
End Sub

Public Function StopFlight(Optional isError As Boolean = False) As Boolean
    StopFlight = True

    'This should never happen:
    If Not info.InFlight Then
        MsgBox "You do not currently have a flight in progress!", vbOKOnly Or vbExclamation, "Error"
        StopFlight = False
        Exit Function
    End If

    'Make sure the pilot wants to stop the flight, but only if not in a flight
    'phase where we have enough info to file a PIREP.
    If Not info.FlightData And Not isError Then
        Dim intResponse As Integer
        If (MsgBox("You have not yet completed this flight!" & vbCrLf & _
            vbCrLf & "The flight is considered complete when you have taken off, landed," & _
            vbCrLf & "taxied to the gate and set the Parking Brake. You will not be able" & _
            vbCrLf & "to file a Flight Report until the flight has been completed and you" & _
            vbCrLf & "are parked at the Gate." & vbCrLf & vbCrLf & _
            "Are you sure you want to end the flight early?", vbYesNo + vbExclamation, _
            "Abort Flight") = vbNo) Then
            StopFlight = False
            Exit Function
        End If
    End If

    'Stop tracking/timing flight.
    tmrPosUpdates.enabled = False
    tmrFlightTime.enabled = False

    'Set some flags and control states.
    If info.FlightData Then
        Dim distance As Integer, distanceD As Integer
        Dim AirportA As Airport
        
        info.FlightPhase = COMPLETE
        
        'Check distance to destination
        distanceD = info.AirportA.DistanceTo(pos.Latitude, pos.Longitude)
        If (distanceD > 10) Then
            Set AirportA = config.GetClosestAirport(pos.Latitude, pos.Longitude)
            distance = AirportA.DistanceTo(pos.Latitude, pos.Longitude)
            
            If (MsgBox("You are " & CStr(distanceD) & " miles away from your Destination." & _
                vbCrLf & "The closest airport (" & CStr(distance) & " miles away) is " & _
                AirportA.name & "." & vbCrLf & vbCrLf & "Update your destination?", _
                vbYesNo Or vbExclamation, "Update Destination") = vbYes) Then
                Set info.AirportA = AirportA
                config.UpdateFlightInfo
            End If
        End If
    Else
        LockFlightInfo True
        If isError Then
            info.FlightPhase = ERROR
        Else
            info.FlightID = 0
            info.FlightPhase = ABORTED
            If Positions.HasData Then
                Positions.Clear
                If Not config.ACARSConnected Then sbMain.Panels(1).Text = "Status: Offline (Cache = 0)"
            End If
        End If
    End If
    
    Set pos = Nothing
    sbMain.Panels(2).Text = "Flight Phase: " & info.PhaseName
    info.InFlight = False

    'Clear the flight ID registry entry, and save the flight data
    If Not isError Then
        DeleteSavedFlight SavedFlightID(info), False
        If (info.FlightPhase = COMPLETE) Then PersistFlightData True
        config.SaveFlightCode ""
    End If

    'Send end_flight request if connected
    If config.ACARSConnected Then
        SendEndFlight
        ReqStack.Send
        Sleep 100
        DoEvents
    ElseIf (info.FlightPhase = COMPLETE) Then
        MsgBox "You have completed a Flight. Please Connect to the server to file a Flight Report.", _
            vbOKOnly Or vbInformation, "Flight Complete"
    End If
    
    'Disable online ATC Tab
    SSTab1.TabEnabled(2) = False
    SSTab1.TabVisible(2) = False

    'Close FSUIPC link.
    FSUIPC_Close
    config.FSUIPCConnected = False

    'Show flight info window if the flight was completed.
    SetButtonMenuStates
    If (info.FlightPhase = COMPLETE) Then
        cmdPIREP.visible = True
        cmdPIREP.enabled = config.ACARSConnected
        cmdStartStopFlight.enabled = False
            
        'If we're connected to the server, ask to file a flight report
        If config.ACARSConnected Then
            If (MsgBox("Do you wish to file the Flight Report?", vbInformation + vbYesNo, _
                "File Flight Report") = vbYes) Then Call cmdPIREP_Click
        End If
    End If
End Function

Sub RestoreFlight()
    Dim oldFlight As SavedFlight
    Dim oldID As String, fName As String
    Dim hasFLT As Boolean, dwResult As Long
    Dim totalWait As Integer
    
    'Get the flight ID
    oldID = SavedFlightID(info)
    
    'Restore the flight
    Set oldFlight = RestoreFlightData(oldID)

    'Check if we have the flight saved
    hasFLT = FileExists(config.FS9Files + "\" + "ACARS Flight " + oldID + ".FLT")
    If Not hasFLT Then
        MsgBox "Cannot find Microsoft Flight Simulator saved flight!", vbCritical + vbOKOnly, _
            "Saved Flight not found"
        Exit Sub
    End If
    
    'Make sure FS9 is running
    While Not IsFSRunning
        If (MsgBox("Microsoft Flight Simulator has not been started.", _
            vbExclamation + vbRetryCancel, "Flight Simulator NotStarted") = _
            vbCancel) Then Exit Sub
    Wend
    
    'Connect to FSUIPC - if FS9 not running, abort
    FSUIPC_Connect
    If Not config.FSUIPCConnected Then Exit Sub
    
    'Make sure FS9 is "ready to fly"
    While Not IsFSReady
        If (MsgBox("Microsoft Flight Simulator is not ""Ready to Fly"". Please" + _
            vbCrLf + "ensure that an aircraft is loaded and ready to fly!", _
            vbExclamation + vbRetryCancel, "Flight Simulator Not Ready") = vbCancel) Then
            FSUIPC_Close
            Exit Sub
        End If
    Wend
    
    'Restore the flight
    fName = "ACARS Flight " + oldID + Chr(0)
    Call FSUIPC_WriteS(&H3F04, Len(fName) + 1, fName, dwResult)
    Call FSUIPC_Write(&H3F00, 2, VarPtr(FLIGHT_LOAD), dwResult)
    If Not FSUIPC_Process(dwResult) Then
        FSUIPC_Close
        MsgBox "Error Restoring Flight", vbError + vbOKOnly, "FSUIPC Error"
        Exit Sub
    ElseIf config.ShowDebug Then
        ShowMessage "Restored Flight via FSUIPC", DEBUGTEXTCOLOR
    End If
    
    'Wait until we are ready to fly
    DoEvents
    Sleep 500
    Do
        Sleep 250
        totalWait = totalWait + 250
    Loop Until ((totalWait > 29500) Or IsFSReady)

    'If we're not ready, then kill things
    If Not IsFSReady Then
        FSUIPC_Close
        MsgBox "Error Restoring Flight", vbError + vbOKOnly, "FSUIPC Error"
        Exit Sub
    End If
    
    'Load aircraft info
    Set acInfo = GetAircraftInfo()
    
    'Save flight phase
    info.InFlight = True
    info.FlightPhase = oldFlight.FlightInfo.FlightPhase
    sbMain.Panels(2).Text = "Phase: " & info.PhaseName

    'Reset button states
    LockFlightInfo False
    cmdStartStopFlight.Caption = "Stop Flight"
    cmdPIREP.visible = True
    cmdPIREP.enabled = False
    tmrPosUpdates.enabled = True
    tmrFlightTime.enabled = True
End Sub

Sub StartFlight()
    Dim dwResult As Long
    Dim distance As Integer
    
    'Kill the autostart timer
    tmrStartCheck.enabled = False

    If info.InFlight Then
        MsgBox "You already have a flight in progress!", vbOKOnly Or vbExclamation, "Error"
        Exit Sub
    End If
    
    'Populate the flight information
    info.CruiseAltitude = frmMain.txtCruiseAlt.Text
    On Error Resume Next
    info.FlightNumber = CInt(frmMain.txtFlightNumber.Text)
    info.FlightLeg = CInt(frmMain.txtLeg.Text)
    On Error GoTo 0

    'Make sure we're not starting a flight without having PIREPped a previous flight.
    If info.FlightData And Not info.PIREPFiled Then
        If MsgBox("You have not submitted a PIREP for your previous flight. If you start a new flight now, the previous flight data will be discarded. Are you sure?", vbYesNo Or vbQuestion, "Error") _
            = vbNo Then Exit Sub
    End If
    
    'Make sure all required flight data has been entered.
    If txtPilotID.Text = "" Then
        MsgBox "Please enter your pilot ID.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        txtPilotID.SetFocus
        Exit Sub
    ElseIf (info.FlightNumber = 0) Then
        MsgBox "Please enter the Flight number.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        txtFlightNumber.SetFocus
        Exit Sub
    ElseIf (info.FlightLeg <= 0) Or (info.FlightLeg > 5) Then
        MsgBox "Please enter the Leg number.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        txtLeg.SetFocus
        Exit Sub
    ElseIf (Len(info.EquipmentType) < 3) Then
        MsgBox "Please select your aircraft type.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        cboEquipment.SetFocus
        Exit Sub
    ElseIf info.CruiseAltitude = "" Then
        MsgBox "Please enter your cruise altitude.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        txtCruiseAlt.SetFocus
        Exit Sub
    ElseIf (info.airportD Is Nothing) Then
        MsgBox "Please enter your departure airport.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        cboAirportD.SetFocus
        Exit Sub
    ElseIf (info.AirportA Is Nothing) Then
        MsgBox "Please enter your destination airport.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        cboAirportA.SetFocus
        Exit Sub
    ElseIf (info.airportD.ICAO = info.AirportA.ICAO) Then
        MsgBox "Your departure (" & info.airportD.ICAO & ") and destination (" & _
            info.AirportA.ICAO & ") airports cannot be the same.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        cboAirportD.SetFocus
        Exit Sub
    ElseIf info.Route = "" Then
        MsgBox "Please enter your route of flight.", vbOKOnly Or vbExclamation, "Cannot Start Flight"
        txtRoute.SetFocus
        Exit Sub
    End If
    
    'Make sure we're rated for the aircraft
    If config.ACARSConnected And Not info.CheckRide And Not config.HasRating(info.EquipmentType) Then
        If (MsgBox("You do not have a " & info.EquipmentType & " aircraft rating. This Flight Report" & _
            vbCrLf & "may not be approved! Do you wish to continue?", vbExclamation + vbYesNo + _
            vbDefaultButton2, "Not Rated in " & info.EquipmentType) = vbNo) Then Exit Sub
    End If
    
    'Attempt to connect to FSUIPC - Make sure the FSUIPC connection succeeded.
    If Not config.FSUIPCConnected Then
        FSUIPC_Connect
        If Not config.FSUIPCConnected Then Exit Sub
    End If
    
    'Check for SB3
    If config.SB3Support Then
        config.SB3Connected = SB3Connected()
        If (SB3Running() And Not config.SB3Connected) Then
            If (MsgBox("Squawkbox 3 is running, but not connected to VATSIM. Do you want to connect?", _
                vbExclamation + vbYesNo, "Squawkbox 3") = vbYes) Then
                FSUIPC_Close
                config.FSUIPCConnected = False
                Exit Sub
            End If
        End If
        
        'If we're running, automatically set the network to VATSIM
        If (config.SB3Connected And (info.Network = "Offline")) Then
            frmMain.cboNetwork.ListIndex = 1
            info.Network = "VATSIM"
        End If
    End If
    
    'Get aircraft information
    Set acInfo = GetAircraftInfo()
    
    'Make sure the aircraft is parked and on the ground.
    Set pos = RecordFlightData(acInfo)
    If (pos Is Nothing) Then
        Exit Sub
    ElseIf (info.FlightID = 0) And ((Not pos.Parked) Or (Not pos.onGround)) Then
        MsgBox "You must be on the ground with the parking brake set in order to start a flight.", _
            vbExclamation Or vbOKOnly, "Cannot Start Flight"
        config.FSUIPCConnected = False
        FSUIPC_Close
        Exit Sub
    ElseIf ((pos.Latitude = 0) And (pos.Longitude = 0)) Then
        MsgBox "FSUIPC does not appear to be working properly. Check your FSUIPC.LOG.", _
            vbOKOnly Or vbCritical, "FSUIPC Position Error"
        config.FSUIPCConnected = False
        FSUIPC_Close
        Exit Sub
    End If
    
    'Make sure we're close to the specified origin airport
    distance = info.airportD.DistanceTo(pos.Latitude, pos.Longitude)
    If (distance > 10) Then
        ShowMessage "Lat=" + CStr(pos.Latitude) + ", Lon=" + CStr(pos.Longitude), _
            ACARSERRORCOLOR
        ShowMessage "APLat=" + CStr(info.airportD.Latitude) + ", APLon=" + _
            CStr(info.airportD.Longitude), DEBUGTEXTCOLOR
        MsgBox "You are " + CStr(distance) + " miles away from " + info.airportD.name + _
            ". You must be within 10 miles to start a flight.", _
            vbExclamation Or vbOKOnly, "Incorrect Departure Airport"
        config.FSUIPCConnected = False
        FSUIPC_Close
        Exit Sub
    End If
    
    'Check if we're connected, but only if the "fly offline"
    'option is turned off. If not, then attempt to connect.
    Dim result As VbMsgBoxResult
    If Not config.ACARSConnected And Not config.FlyOffline Then
        ToggleACARSConnection

        'Wait for connection results.
        Do
            DoEvents
            Sleep 75
            If config.ACARSConnected Then Exit Do
        Loop While True

        'If the connection failed, prompt to fly offline.
        If Not config.ACARSConnected Then
            result = MsgBox("The connection to the ACARS server failed. Do you wish to fly offline?", vbYesNo Or vbQuestion, "Connection Error")
            If result = vbNo Then Exit Sub
        End If
    End If
    
    'Start the flight
    info.StartFlight config.ACARSConnected
    InitFuelCheck

    'If we're connected to the ACARS server, send a flight info message.
    If config.ACARSConnected Then
        info.InfoReqID = SendFlightInfo(info)
        ReqStack.Send
        
        'Wait for the ACK and the Flight ID
        If Not WaitForACK(info.InfoReqID, 3500) Then
            MsgBox "Time out waiting for Flight ID", vbOKOnly + vbCritical, "Time Out"
            Exit Sub
        End If
    End If
    
    'Write to the gauge
    GAUGE_SetPhase info.FlightPhase, config.ACARSConnected
    GAUGE_SetInfo info, frmMain.txtPilotID.Text

    'Start timing/tracking flight
    tmrFlightTime.enabled = True
    tmrPosUpdates.enabled = True
    
    'Update status bar.
    sbMain.Panels(2).Text = "Flight Phase: " & info.PhaseName
    
    'Save the flight status
    config.SaveFlightCode SavedFlightID(info)

    'Set some flags, variables, and control states.
    SetButtonMenuStates
    LockFlightInfo False
    cmdPIREP.visible = True
    cmdPIREP.enabled = False
End Sub

Private Sub mnuSystrayRestore_Click()
    WindowState = vbNormal
    Show
    TaskBarShow frmMain.hWnd
    RemoveIcon
End Sub

Private Sub tmrFlightTime_Timer()
    Dim fTime As Date

    If (pos Is Nothing) Then Exit Sub
    
    'Update flight time if not pasued/slewing
    If Not pos.Paused And Not pos.Slewing Then
        fTime = info.UpdateFlightTime(pos.simRate / 256, tmrFlightTime.interval)
        sbMain.Panels(4).Text = "Flight Time: " + Format(fTime, "hh:mm:ss")
    End If
        
    'Force the sim rate to be within the allowed range.
    CheckSimRate MINTIMECOMPRESSION, MAXTIMECOMPRESSION
End Sub

Private Sub tmrPosUpdates_Timer()
    Static LastPosUpdate As Date
    Static LastATCUpdate As Date
    Static LastFlightSave As Date
    Static PauseStatus As Boolean
    
    Dim isPaused As Boolean
    Dim CurrentDate As Date
    
    'Get position data
    Set pos = RecordFlightData(acInfo)
    If (pos Is Nothing) Then
        tmrPosUpdates.enabled = False
        Exit Sub
    End If
    
    'Display if we're paused/replaying
    isPaused = (pos.Paused Or pos.Slewing)
    If isPaused Then
        sbMain.Panels(2).Text = "Phase: Paused/Slewing"
        PauseStatus = True
    ElseIf PauseStatus Then
        sbMain.Panels(2).Text = "Phase: " & info.PhaseName
        PauseStatus = False
    End If

    'Check if the flight phase has changed
    If Not isPaused Then
        If PhaseChanged(pos) Then sbMain.Panels(2).Text = "Phase: " & info.PhaseName
    End If

    'Check if it's time to send a position update.
    CurrentDate = Now
    If (IsEmpty(LastPosUpdate)) Or (DateDiff("s", LastPosUpdate, CurrentDate) > config.PositionInterval) Then
        LastPosUpdate = CurrentDate
        sbMain.Panels(3).Text = "Last Position Report: " & Format(LastPosUpdate, "hh:mm:ss")
        
        'Send data to the server. Otherwise save it.
        If config.ACARSConnected And (info.FlightID > 0) Then
            SendPosition pos
        ElseIf IsDate(info.StartTime.UTCTime) And Not isPaused Then
            Positions.AddPosition pos
            sbMain.Panels(1).Text = "Status: Offline (Cache = " + Format(Positions.Size, "#,##0") + ")"
            If config.ShowDebug Then ShowMessage "Position Cache = " + Format(Positions.Size, "#,##0"), DEBUGTEXTCOLOR
        End If
    End If
    
    'Check if it's time to request ATC info
    If (config.ACARSConnected And (Not isPaused) And (info.FlightID > 0) And (info.Network <> "Offline")) Then
        If (IsEmpty(LastATCUpdate) Or (DateDiff("s", LastATCUpdate, CurrentDate) > 190)) Then
            LastATCUpdate = CurrentDate
            RequestATCInfo info.Network
        End If
    End If
    
    'Check if it's time to save the flight
    If (config.IsFS9 And (Not isPaused)) Then
        If (IsEmpty(LastFlightSave) Or (DateDiff("s", LastFlightSave, CurrentDate) > 60)) Then
            LastFlightSave = CurrentDate
            PersistFlightData
            SaveFlight
        End If
    End If
    
    'If we have request data waiting, send it
    If ReqStack.HasData Then ReqStack.Send

    'Calculate the update interval based on our phase/ground speed
    Dim newInterval As Integer
    newInterval = UpdatePositionInterval()
    If (config.PositionInterval <> newInterval) Then
        config.PositionInterval = newInterval
        If config.ShowDebug Then ShowMessage "Position Interval set to " + CStr(newInterval) + "s", DEBUGTEXTCOLOR
    End If
End Sub

Private Sub ToggleACARSConnection(Optional silent As Boolean = False)
    If Not config.ACARSConnected Then
        'Validate credentials
        If (Len(frmMain.txtPilotID.Text) < 4) Then
            MsgBox "Please enter your User ID.", vbCritical + vbOKOnly, "Invalid Credentials"
            frmMain.txtPilotID.SetFocus
            Exit Sub
        ElseIf (Len(frmMain.txtPassword.Text) < 2) Then
            MsgBox "Please enter your Password.", vbCritical + vbOKOnly, "Invalid Credentials"
            frmMain.txtPassword.SetFocus
            Exit Sub
        End If
    
        cmdConnectDisconnect.enabled = False
        On Error GoTo EH
        wsckMain.Close
        wsckMain.RemoteHost = config.ACARSHost
        wsckMain.RemotePort = config.ACARSPort
        config.SeenHELO = False
        wsckMain.Connect
        frmMain.cmdPIREP.visible = info.FlightData
        frmMain.cmdPIREP.enabled = info.FlightData
        LockUserInfo False
    Else
        cmdConnectDisconnect.enabled = False
        
        'Confirm disconnect if silent isn't set
        If Not silent Then
            If Not ConfirmDisconnect Then
                cmdConnectDisconnect.enabled = True
                Exit Sub
            End If
        End If
        
        CloseACARSConnection True
        cmdConnectDisconnect.enabled = True
        info.Offline = True
        frmMain.cmdPIREP.visible = False
        LockUserInfo True
    End If
    
ExitSub:
    Exit Sub
    
EH:
    cmdConnectDisconnect.enabled = True
    If ((err.Number <> 40060) And (err.Number <> 10053)) Then
        MsgBox "The following error occurred: " & err.Description & " (" & err.Number & ")", vbOKOnly Or vbCritical, "ToggleACARSConnection.Error"
    End If
    
    Resume ExitSub
End Sub

Private Sub tmrPing_Timer()
    If Not config.ACARSConnected Then
        tmrPing.enabled = False
        Exit Sub
    End If
    
    'Send a ping if it's been too long
    If (DateDiff("s", ReqStack.LastUse, Now) > config.PingInterval) Then
        SendPing
        ReqStack.Send
        DoEvents
    End If
End Sub

Private Sub tmrStartCheck_Timer()
    Dim fdrData As PositionData
    Static acInfo As AircraftInfo

    'Check if FS9 is running and FSUIPC is connected
    If Not IsFSReady Then Exit Sub
    
    'Instantiate aircraft info
    If (acInfo Is Nothing) Then
        Set acInfo = GetAircraftInfo()
        If (acInfo Is Nothing) Then Exit Sub
    End If
    
    'Get flight data
    Set fdrData = RecordFlightData(acInfo)
    If ((Not fdrData.Parked) And (Abs(fdrData.GroundSpeed) <= 3)) Then
        Dim pushbackOff As Integer
        Dim brakesOn As Integer
        Dim dwResult As Long
        
        ShowMessage "Parkbrake = " + CStr(fdrData.Parked), DEBUGTEXTCOLOR
        ShowMessage "GroundSpeed = " + CStr(fdrData.GroundSpeed), DEBUGTEXTCOLOR
    
        'turn off pushback/turn on brakes gently
        pushbackOff = 3
        brakesOn = 10240
        Call FSUIPC_Write(&H31F0, 2, VarPtr(pushbackOff), dwResult)
        Call FSUIPC_Write(&HBC4, 2, VarPtr(brakesOn), dwResult)
        Call FSUIPC_Write(&HBC6, 2, VarPtr(brakesOn), dwResult)
        If Not FSUIPC_Process(dwResult) Then
            ShowMessage "Error canceling pushback/start", ACARSERRORCOLOR
            If config.ShowDebug Then ShowMessage "Disabling Flight Start Timer", DEBUGTEXTCOLOR
            tmrStartCheck.enabled = False
            Exit Sub
        End If
        
        'Wait a teensy bit
        Sleep 275
        
        'turn on parking brake
        brakesOn = 0
        pushbackOff = 32767
        Call FSUIPC_Write(&HBC4, 2, VarPtr(brakesOn), dwResult)
        Call FSUIPC_Write(&HBC6, 2, VarPtr(brakesOn), dwResult)
        Call FSUIPC_Write(&HBC8, 2, VarPtr(pushbackOff), dwResult)
        
        'Write a message - this will write the calls above
        ShowFSMessage "Pushback Canceled - No Flight Started", True, 15
        PlaySoundFile "notify_error.wav"
        If config.ShowDebug Then ShowMessage "Disabling Flight Start Timer", DEBUGTEXTCOLOR
        tmrStartCheck.enabled = False
    ElseIf (Abs(fdrData.GroundSpeed) > 3) Then
        tmrStartCheck.enabled = False
        ShowMessage "Disabling Flight Start Timer", ACARSTEXTCOLOR
    End If
End Sub

Private Sub txtAirportA_Change()
    LimitLength txtAirportA, 4, True
End Sub

Private Sub txtAirportA_GotFocus()
    SelectField txtAirportA
End Sub

Private Sub txtAirportA_LostFocus()
    Dim ap As Airport
    
    Set ap = config.GetAirport(txtAirportA.Text)
    If Not (ap Is Nothing) Then
        Set info.AirportA = ap
        config.UpdateFlightInfo
    Else
        txtAirportA.Text = ""
    End If
End Sub

Private Sub txtAirportD_Change()
    LimitLength txtAirportD, 4, True
End Sub

Private Sub txtAirportD_GotFocus()
    SelectField txtAirportD
End Sub

Private Sub txtAirportD_LostFocus()
    Dim ap As Airport
    
    Set ap = config.GetAirport(txtAirportD.Text)
    If Not (ap Is Nothing) Then
        Set info.airportD = ap
        config.UpdateFlightInfo
    Else
        txtAirportD.Text = ""
    End If
End Sub

Private Sub txtAirportL_Change()
    LimitLength txtAirportL, 4, True
End Sub

Private Sub txtAirportL_GotFocus()
    SelectField txtAirportL
End Sub

Private Sub txtAirportL_LostFocus()
    Dim ap As Airport
    
    Set ap = config.GetAirport(txtAirportL.Text)
    If Not (ap Is Nothing) Then
        Set info.AirportL = ap
        config.UpdateFlightInfo
    Else
        txtAirportL.Text = ""
    End If
End Sub

Private Sub txtCmd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        KeyCode = 0
        If blnFirstCmd Then Exit Sub
        If intCmdHistIndex > UBound(CmdHistory) Then
            strCmdBuffer = txtCmd.Text
            intCmdBufferSel = txtCmd.SelStart
        End If
        intCmdHistIndex = intCmdHistIndex - 1
        If intCmdHistIndex = -1 Then
            intCmdHistIndex = 0
            Exit Sub
        Else
            txtCmd.Text = CmdHistory(intCmdHistIndex)
            txtCmd.SelLength = 0
            txtCmd.SelStart = Len(txtCmd.Text)
        End If
    ElseIf KeyCode = 40 Then
        KeyCode = 0
        If blnFirstCmd Then Exit Sub
        If intCmdHistIndex > UBound(CmdHistory) Then Exit Sub
        intCmdHistIndex = intCmdHistIndex + 1
        If intCmdHistIndex = (UBound(CmdHistory) + 1) Then
            txtCmd.Text = strCmdBuffer
            txtCmd.SelLength = 0
            txtCmd.SelStart = intCmdBufferSel
        Else
            txtCmd.Text = CmdHistory(intCmdHistIndex)
            txtCmd.SelLength = 0
            txtCmd.SelStart = Len(txtCmd.Text)
        End If
    End If
End Sub

Private Sub txtCmd_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        KeyAscii = 0
        If (txtCmd.Text <> "") Then
            If Not blnFirstCmd Then
                ReDim Preserve CmdHistory(UBound(CmdHistory) + 1) As String
            Else
                ReDim CmdHistory(0) As String
                blnFirstCmd = False
            End If
            CmdHistory(UBound(CmdHistory)) = txtCmd.Text
            ProcessUserInput CStr(txtCmd.Text)
            txtCmd.Text = ""
            strCmdBuffer = ""
            intCmdBufferSel = 0
            intCmdHistIndex = UBound(CmdHistory) + 1
        End If
    End If
End Sub

Private Sub txtCruiseAlt_Change()
    LimitLength txtCruiseAlt, 5, True
End Sub

Private Sub txtCruiseAlt_GotFocus()
    SelectField txtCruiseAlt
End Sub

Private Sub txtFlightNumber_LostFocus()
    On Error GoTo ResetNumber
    LimitLength txtFlightNumber, 4, True
    info.FlightNumber = CInt(txtFlightNumber.Text)
    
ExitSub:
    Exit Sub
    
ResetNumber:
    txtFlightNumber.Text = "001"
    info.FlightNumber = 1
    Resume ExitSub
    
End Sub

Private Sub txtLeg_Change()
    LimitLength txtLeg, 1
    If (InStr(1, "123456", txtLeg.Text) < 1) Or (Len(txtLeg.Text) = 0) Then txtLeg.Text = 1
    info.FlightLeg = CInt(txtLeg.Text)
End Sub

Private Sub txtPilotID_Change()
    LimitLength txtPilotID, 8, True
End Sub

Private Sub txtRemarks_Change()
    info.Remarks = txtRemarks.Text
End Sub

Private Sub txtRoute_LostFocus()
    info.Route = UCase(txtRoute.Text)
End Sub

Private Sub wsckMain_Close()
    Dim wasConnected As Boolean

    wasConnected = config.ACARSConnected

    CloseACARSConnection False
    If config.FSUIPCConnected Then GAUGE_Disconnect
    If wasConnected Then
        PlaySoundFile "notify_error.wav"
        ShowMessage "ACARS connection closed by server!", ACARSERRORCOLOR
        If info.InFlight Then info.Offline = True
    End If
End Sub

Private Sub wsckMain_Connect()
    config.ACARSConnected = True
    tmrPing.enabled = True
    cmdConnectDisconnect.Caption = "Disconnect"
    mnuConnect.Caption = "Disconnect"
    sbMain.Panels(1).Text = "Status: Connected"
    
    'Log in
    info.AuthReqID = SendCredentials(frmMain.txtPilotID.Text, frmMain.txtPassword.Text)
    ReqStack.Send
    
    'Wait for an ACK
    If Not WaitForACK(info.AuthReqID, 9750) Then
        If (info.AuthReqID <> 0) Then
            info.AuthReqID = 0 'Discard the ACK if it comes back
            MsgBox "ACARS Authentication timed out!", vbOKOnly + vbCritical, "Timed Out"
        End If
        
        If config.ACARSConnected Then ToggleACARSConnection True
    Else
        cmdConnectDisconnect.enabled = True
        sbMain.Panels(1).Text = "Status: Connected to ACARS server"
    End If
End Sub

Private Sub wsckMain_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    wsckMain.GetData strData, vbString
    ProcessServerData strData
End Sub

Private Sub wsckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'If we get a 40006 or 10053 then toggle the ACARS connection
    PlaySoundFile "notify_error.wav"
    If ((Number = 40006) Or (Number = 10053)) Then
        CloseACARSConnection False
        If info.InFlight Then info.Offline = True
        ShowMessage "Lost Connection to ACARS Server", ACARSERRORCOLOR
    ElseIf (Number = 10061) Then
        CloseACARSConnection False
        If info.InFlight Then info.Offline = True
        ShowMessage "Cannot connect to ACARS Server at " + config.ACARSHost + ".", ACARSERRORCOLOR
        ShowMessage "Make sure your firewall is not blocking Port " + CStr(config.ACARSPort) + ".", _
            ACARSERRORCOLOR
    Else
        MsgBox "The following error occurred: " & Description & " (" & Number & ")", vbOKOnly Or vbCritical, "wsckMain.Error"
        CloseACARSConnection False
    End If
End Sub

Private Function ConfirmDisconnect() As Boolean
    ConfirmDisconnect = (MsgBox("Are you sure you want to disconnect?", vbYesNo Or vbQuestion, "Confirm") = vbYes)
End Function

Public Sub CloseACARSConnection(Optional blnSendQuit As Boolean = False)
    config.ACARSConnected = False
    
    'Disconnect
    On Error Resume Next
    wsckMain.Close
    On Error GoTo 0
    DoEvents
    
    'Update status
    tmrPing.enabled = False
    sbMain.Panels(1).Text = "Status: Offline"
    mnuConnect.Caption = "Connect"
    cmdConnectDisconnect.Caption = "Connect"
    cmdConnectDisconnect.enabled = True
    mnuOptionsFlyOffline.enabled = True
    
    'Determine if we are dispatch
    If (config.HasRole("HR") Or config.HasRole("Dispatch")) Then
        users.DeletePilot frmMain.txtPilotID
        If Not users.DispatchOnline Then ShowMessage "ACARS Messaging Restrictions restored by your logout", _
            ACARSTEXTCOLOR
    End If
    
    txtPilotID.enabled = True
    txtPassword.enabled = True
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabVisible(2) = False
    SSTab1.Tab = 0
End Sub

Public Sub ProcessServerData(strData As String)
    Static strInputBuffer As String
    Dim intPos As Long

    'Ignore initial HELLO string.
    If Not config.SeenHELO Then
        config.SeenHELO = (InStr(1, strData, "HELLO", vbTextCompare) > 0)
        Exit Sub
    End If

    'Add new data to input buffer.
    strInputBuffer = strInputBuffer & strData
    Do
        'Look for </ACARSResponse> signifying end of XML block.
        intPos = 0
        intPos = InStr(1, strInputBuffer, "</" & XMLRESPONSEROOT & ">", vbTextCompare)
        If intPos > 0 Then
            'If found, chop off everything up to and including </ACARSResponse>
            Dim Parts As Variant
            Parts = Split(strInputBuffer, "</" & XMLRESPONSEROOT & ">", 2, vbTextCompare)
            
            'Send XML message to ProcessServerMessage and leave rest in buffer
            ProcessMessage CStr(Parts(0)) & "</" & XMLRESPONSEROOT & ">"
            strInputBuffer = CStr(Parts(1))
        End If
    Loop Until intPos = 0
End Sub

Private Sub ProcessUserInput(strInput As String)
    If (InStr(strInput, ".") <> 1) And config.ACARSConnected Then
        SendChat strInput
        ReqStack.Send
    Else
        ProcessUserCmd strInput
    End If
End Sub

Private Sub ProcessUserCmd(strInput As String)
    Static MsgCount As Integer
    Dim cmdName As String
    Dim aryParts As Variant
    
    strInput = Mid(strInput, 2, Len(strInput) - 1)
    aryParts = Split(strInput, " ", 2, vbTextCompare)
    cmdName = aryParts(0)

    'Process command accordingly.
    Select Case cmdName
        Case "msg"
            Dim DispatchOnline As Boolean
        
            'Make sure we're connected
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            ElseIf UBound(aryParts) < 1 Then
                ShowMessage "No message specified!", ACARSERRORCOLOR
                Exit Sub
            End If
                
            'Make sure a message was specified.
            aryParts = Split(aryParts(1), " ", 2, vbTextCompare)
            If UBound(aryParts) < 1 Then
                ShowMessage "No message specified!", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'If we're not flying, then check our status
            If config.NoMessages Then
                ShowMessage "ACARS Messaging Disabled", ACARSERRORCOLOR
                PlaySoundFile "notify.error_wav"
                Exit Sub
            ElseIf Not info.InFlight And Not config.IsUnrestricted And Not users.DispatchOnline Then
                If (MsgCount >= MAX_TEXT_MESSAGES) Then
                    ShowMessage "You are limited to sending " + CStr(MAX_TEXT_MESSAGES) + _
                    " ACARS messages when not flying and Dispatch is not online.", ACARSERRORCOLOR
                    PlaySoundFile "notify.error_wav"
                    Exit Sub
                Else
                    MsgCount = MsgCount + 1
                    If config.ShowDebug Then ShowMessage "Sending message " + CStr(MsgCount) + _
                        " of " + CStr(MAX_TEXT_MESSAGES) + " without Dispatch", DEBUGTEXTCOLOR
                End If
            End If

            'Process the chat message.
            SendChat CStr(aryParts(1)), CStr(aryParts(0))
            ReqStack.Send
            
        Case "pvtvoice"
            'Make sure we're online
            If Not config.SB3Support Then Exit Sub
            If Not config.SB3Connected Then
                ShowMessage "Squawkbox 3 not connected to VATSIM Server", ACARSERRORCOLOR
                Exit Sub
            ElseIf (config.PrivateVoiceURL = "") Then
                ShowMessage "No Private Voice URL", ACARSERRORCOLOR
                Exit Sub
            End If
            
            SB3PrivateVoice config.PrivateVoiceURL
            ShowMessage "Private Voice Channel set to " + config.PrivateVoiceURL, ACARSTEXTCOLOR
            
        Case "update"
            'Make sure we're connected
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            End If

            RequestAirlines
            RequestEquipment
            RequestAirports
            ReqStack.Send

        Case "info"
            'Make sure we're connected
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            End If

            'Make sure a pilot ID was specified.
            If UBound(aryParts) < 1 Then
                ShowMessage "No pilot ID specified!", ACARSERRORCOLOR
                Exit Sub
            End If

            'Send the command.
            RequestPilotInfo CStr(aryParts(1))
            ReqStack.Send
            
        Case "charts"
            'Make sure an airport was specified
            If UBound(aryParts) < 1 Then
                ShowMessage "No Airport specified", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Make sure we're connected and a flight was started
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            End If
            
            RequestCharts CStr(aryParts(1))
            ReqStack.Send
            
        Case "nav1", "nav2"
            'Make sure a navaid was specified
            If UBound(aryParts) < 1 Then
                ShowMessage "No VOR/ILS specified", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Make sure we're connected and a flight was started
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            ElseIf Not info.InFlight Then
                ShowMessage "Flight Not Started", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Make sure a runway was specified
            aryParts = Split(aryParts(1), " ", 2, vbTextCompare)
            If UBound(aryParts) < 1 Then
                ShowMessage "No Heading/Runway specified", ACARSERRORCOLOR
                Exit Sub
            End If

            RequestNavaidInfo CStr(aryParts(0)), CStr(aryParts(1)), cmdName
            ReqStack.Send
            
        Case "runway"
            'Make sure a navaid was specified
            If UBound(aryParts) < 1 Then
                ShowMessage "No Airport specified", ACARSERRORCOLOR
                Exit Sub
            End If
            
            aryParts = Split(aryParts(1), " ")
            If (UBound(aryParts) < 1) Then
                ShowMessage "No Runway specified", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Make sure we're connected
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            ElseIf Not info.InFlight Then
                ShowMessage "Flight Not Started", ACARSERRORCOLOR
                Exit Sub
            End If

            RequestRunwayInfo CStr(aryParts(0)), CStr(aryParts(1))
            ReqStack.Send
            
        Case "com1"
            Dim freq As String
        
            'Make sure a frequency was specified
            If UBound(aryParts) < 1 Then
                ShowMessage "No Frequency specified", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Tune the COM1 radio to the frequency
            freq = CStr(aryParts(1))
            SetCOM1 freq
            ShowMessage "COM1 Radio set to " + freq, ACARSTEXTCOLOR
            
        Case "kick", "ban"
            'Check our access Make sure a Pilot ID was specified
            If Not config.HasRole("Admin") Then
                ShowMessage "Insufficient Access", ACARSERRORCOLOR
                Exit Sub
            ElseIf UBound(aryParts) < 1 Then
                ShowMessage "No Pilot ID specified", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Kick the user
            SendKick CStr(aryParts(1)), (cmdName = "ban")
            ReqStack.Send
            
        Case "draft"
            RequestDraftPIREPs
            ReqStack.Send
            
        Case "atc"
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            ElseIf (info.Network = "Offline") Then
                ShowMessage "Not Connected to ATC Network", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Request ATC info
            RequestATCInfo info.Network
            ReqStack.Send
            
        Case "progress"
            Dim AirportC As Airport
            Dim distanceD As Integer, distanceA As Integer, distanceL As Integer
            Dim distanceC As Integer, timeToC As Integer
            Dim timeToA As Long, timeToL As Long
        
            'Make sure we're in flight
            If ((info.FlightPhase <> AIRBORNE) Or (pos Is Nothing)) Then
                ShowMessage "Not currently airborne", ACARSERRORCOLOR
                Exit Sub
            End If
            
            'Get closest airport
            Set AirportC = config.GetClosestAirport(pos.Latitude, pos.Longitude)
            
            'Calculate distance from source/destination
            distanceD = info.airportD.DistanceTo(pos.Latitude, pos.Longitude)
            distanceA = info.AirportA.DistanceTo(pos.Latitude, pos.Longitude)
            distanceC = AirportC.DistanceTo(pos.Latitude, pos.Longitude)
            timeToA = (distanceA * 60&) / pos.GroundSpeed    'Minutes
            timeToC = (distanceC * 60&) / pos.GroundSpeed
            
            'Display distance
            ShowMessage "Distance from " & info.airportD.name & ": " & CStr(distanceD) & " miles", ACARSTEXTCOLOR
            ShowMessage "Distance to " & info.AirportA.name & ": " & CStr(distanceA) & " miles (" & _
                CStr(timeToA) & " minutes)", ACARSTEXTCOLOR
            ShowMessage "Distance to " & AirportC.name & ": " & CStr(distanceC) & " miles (" & _
                CStr(timeToC) & " minutes)", ACARSTEXTCOLOR
            
            'Display alternate info
            If Not (info.AirportL Is Nothing) Then
                distanceL = info.AirportL.DistanceTo(pos.Latitude, pos.Longitude)
                timeToL = (distanceL * 60&) / pos.GroundSpeed
                ShowMessage "Distance to " & info.AirportL.name & ": " & CStr(distanceL) & _
                    " miles (" & CStr(timeToL) & " minutes)", ACARSTEXTCOLOR
            End If
        
        Case "help"
            ShowMessage "Delta Virtual Airlines ACARS Help", ACARSTEXTCOLOR
            
            'Show commmands requiring ACARS connection
            If config.ACARSConnected Then
                'Display msg if we can send messages
                If config.IsUnrestricted Then
                    ShowMessage ".msg [userid] <msg> - Sends a message to another user", ACARSTEXTCOLOR
                ElseIf Not config.NoMessages And users.DispatchOnline Then
                    ShowMessage ".msg [userid] <msg> - Sends a message to another user", ACARSTEXTCOLOR
                End If
                
                ShowMessage ".runway <airport> <runway> - Loads runway data/tunes ILS (if present)", ACARSTEXTCOLOR
                ShowMessage ".charts <airport> - Loads approach charts", ACARSTEXTCOLOR
                ShowMessage ".update - Update Aircraft/Airport choices", ACARSTEXTCOLOR
                If info.InFlight Then
                    ShowMessage ".nav1 <vor> <heading> - Tunes NAV1 radio to a VOR", ACARSTEXTCOLOR
                    ShowMessage ".nav2 <vor> <heading> - Tunes NAV2 radio to a VOR", ACARSTEXTCOLOR
                Else
                    ShowMessage ".draft - Display draft Flight Reports", ACARSTEXTCOLOR
                End If
                
                If (info.Network <> "Offline") Then ShowMessage ".atc - Request " & info.Network & _
                    " ATC Information", ACARSTEXTCOLOR
            End If
            
            'Show command requiring active flight
            If info.InFlight Then
                ShowMessage ".com1 <frequency> - Tunes COM1 radio to a frequency", ACARSTEXTCOLOR
                ShowMessage ".progress - Show Flight Progress", ACARSTEXTCOLOR
            End If
            
            'Show commands requiring SB3
            If config.SB3Connected Then _
                ShowMessage ".pvtvoice - Tunes to DVA Private Voice Channel", ACARSTEXTCOLOR
            
            'Show admin commands
            If config.HasRole("Admin") Then
                ShowMessage ".kick <pilot> - Remove Pilot", ACARSTEXTCOLOR
                ShowMessage ".ban <pilot> - Remove Pilot and prevent logins", ACARSTEXTCOLOR
            End If
            
            ShowMessage ".help - Display this help screen", ACARSTEXTCOLOR
            
        Case Else
            ShowMessage "Unknown command: " & aryParts(0), ACARSERRORCOLOR

    End Select
End Sub

Public Sub LockFlightInfo(IsEditable As Boolean)
    txtFlightNumber.enabled = IsEditable
    txtLeg.enabled = IsEditable
    txtCruiseAlt.enabled = IsEditable
    cboEquipment.enabled = IsEditable
    cboAirline.enabled = IsEditable
    cboNetwork.enabled = IsEditable
    cboAirportD.enabled = IsEditable
    txtAirportD.enabled = IsEditable
    cboAirportA.enabled = IsEditable
    txtAirportA.enabled = IsEditable
    cboAirportL.enabled = IsEditable
    txtAirportL.enabled = IsEditable
    chkCheckRide.enabled = IsEditable
End Sub

Private Sub LockUserInfo(IsEditable As Boolean)
    txtPilotID.enabled = IsEditable
    txtPassword.enabled = IsEditable
End Sub

Private Sub LimitLength(txt As TextBox, ByVal maxLen As Integer, Optional doUpperCase As Boolean = False)
    If (Len(txt.Text) > maxLen) Then txt.Text = Left(txt.Text, maxLen)
    If (doUpperCase And (UCase(txt.Text) <> txt.Text)) Then
        txt.Text = UCase(txt.Text)
        txt.SelStart = Len(txt.Text)
    End If
End Sub

Private Sub SelectField(txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub

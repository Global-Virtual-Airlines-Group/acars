VERSION 5.00
Begin VB.Form frmDraftPIREP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Draft Flight Reports"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "frmDraftPIREP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Flight Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   3450
      Width           =   2895
   End
   Begin VB.Frame frameFlightInfo 
      Caption         =   "Flight Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1555
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label AirportA 
         BackStyle       =   0  'Transparent
         Caption         =   "Airport Name (XXXX)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1220
         Width           =   2535
      End
      Begin VB.Label airportD 
         BackStyle       =   0  'Transparent
         Caption         =   "Airport Name (XXXX)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   930
         Width           =   2535
      End
      Begin VB.Label eqType 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label flightCode 
         BackStyle       =   0  'Transparent
         Caption         =   "DVAXXXX Leg X"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label lblAirportA 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Arriving at"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1220
         Width           =   1095
      End
      Begin VB.Label lblAirportD 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Flying from"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label lblEquipment 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   585
         Width           =   1095
      End
      Begin VB.Label lblFlightCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Flight Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.ListBox lstPIREP 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5175
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select a Flight Report to load from the list below."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmDraftPIREP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public flights As Variant

Private Sub cmdLoad_Click()
    Dim pirep As FlightReport
    
    'Load the flight data
    Set pirep = flights(lstPIREP.ListIndex)
    Set info.Airline = pirep.Airline
    info.FlightNumber = pirep.FlightNumber
    info.FlightLeg = pirep.Leg
    info.EquipmentType = pirep.EquipmentType
    Set info.airportD = pirep.airportD
    Set info.AirportA = pirep.AirportA
    info.Remarks = pirep.Remarks
    info.Network = pirep.Network
    config.UpdateFlightInfo
    
    'Unload this form
    Unload frmDraftPIREP
End Sub

Private Sub Form_Load()
    flights = Array()
End Sub

Private Sub lstPIREP_Click()
    Dim pirep As FlightReport

    'Load the flight data
    Set pirep = flights(lstPIREP.ListIndex)
    flightCode.Caption = pirep.Airline.code + CStr(pirep.FlightNumber) + _
        " Leg " + CStr(pirep.Leg)
    eqType.Caption = pirep.EquipmentType
    airportD.Caption = pirep.airportD.name + " (" + pirep.airportD.ICAO + ")"
    AirportA.Caption = pirep.AirportA.name + " (" + pirep.AirportA.ICAO + ")"
    
    'Display stuff
    frameFlightInfo.visible = True
    cmdLoad.enabled = True
End Sub

Public Sub AddFlight(f As FlightReport)
    ReDim Preserve flights(UBound(flights) + 1)
    Set flights(UBound(flights)) = f
End Sub

Public Function Size() As Integer
    Size = UBound(flights) + 1
End Function

Public Sub Update()
    Dim x As Integer
    Dim pirep As FlightReport
    
    lstPIREP.Clear
    For x = 0 To UBound(flights)
        Set pirep = flights(x)
        lstPIREP.AddItem pirep.Description
    Next
        
    lstPIREP.enabled = True
End Sub

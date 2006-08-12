VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFuel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aircraft Fuel Calculator / Loader"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmLoader 
      Caption         =   "Fuel Tank Loader"
      Height          =   2980
      Left            =   120
      TabIndex        =   1
      Top             =   4050
      Visible         =   0   'False
      Width           =   6855
      Begin VB.TextBox txtMaxFuel 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4920
         TabIndex        =   59
         Top             =   1920
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.CommandButton cmdFuelLoad 
         Caption         =   "Load Aircraft Fuel Tanks"
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtTotalFuel 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   43
         Top             =   1920
         Visible         =   0   'False
         Width           =   700
      End
      Begin MSComctlLib.Slider sldLeftTip 
         Height          =   1335
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldLeftAUX 
         Height          =   1335
         Left            =   1000
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldLeftMain 
         Height          =   1335
         Left            =   1720
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldCenter 
         Height          =   1335
         Left            =   2440
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldRightMain 
         Height          =   1335
         Left            =   3160
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldRightAUX 
         Height          =   1335
         Left            =   3900
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldRightTip 
         Height          =   1335
         Left            =   4620
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldCenter2 
         Height          =   1335
         Left            =   5540
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin MSComctlLib.Slider sldCenter3 
         Height          =   1335
         Left            =   6220
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   2355
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   100
         SelectRange     =   -1  'True
         TextPosition    =   1
      End
      Begin VB.Label lblMaxWeightWarning 
         Alignment       =   2  'Center
         Caption         =   "FUEL LOAD EXCEEDS MAXIMUM GROSS WEIGHT - REDUCE PAYLOAD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   2280
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label lblMaxFuelPounds 
         Caption         =   "pounds"
         Height          =   255
         Left            =   5760
         TabIndex        =   60
         Top             =   1950
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblMaxFuel 
         Alignment       =   1  'Right Justify
         Caption         =   "FUEL CAPACITY"
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
         Left            =   3480
         TabIndex        =   58
         Top             =   1950
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTotalFuel 
         Alignment       =   1  'Right Justify
         Caption         =   "FUEL TO LOAD"
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
         Left            =   120
         TabIndex        =   45
         Top             =   1950
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblTotalFuelPounds 
         Caption         =   "pounds"
         Height          =   255
         Left            =   2580
         TabIndex        =   44
         Top             =   1950
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCenter3 
         Alignment       =   2  'Center
         Caption         =   "Center 3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6080
         TabIndex        =   42
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblCenter2 
         Alignment       =   2  'Center
         Caption         =   "Center 2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5380
         TabIndex        =   41
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblRightTip 
         Alignment       =   2  'Center
         Caption         =   "Right Tip"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4520
         TabIndex        =   40
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblRightAUX 
         Alignment       =   2  'Center
         Caption         =   "Right AUX"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3750
         TabIndex        =   39
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblRightMain 
         Alignment       =   2  'Center
         Caption         =   "Right Main"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   38
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblCenter 
         Alignment       =   2  'Center
         Caption         =   "Center"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2340
         TabIndex        =   37
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblLeftMain 
         Alignment       =   2  'Center
         Caption         =   "Left Main"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1590
         TabIndex        =   36
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblLeftAUX 
         Alignment       =   2  'Center
         Caption         =   "Left AUX"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   18
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblLeftTip 
         Alignment       =   2  'Center
         Caption         =   "Left Tip"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame frmCalculator 
      Caption         =   "Fuel Load Calculator"
      Height          =   3820
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtFuelRequired 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1095
      End
      Begin VB.TextBox txtProfileName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1290
         Width           =   4695
      End
      Begin VB.TextBox txtMaxWeight 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   2025
         Width           =   975
      End
      Begin VB.TextBox txtFuelFlow 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   2025
         Width           =   735
      End
      Begin VB.ComboBox cboProfile 
         Height          =   315
         ItemData        =   "frmFuel.frx":0000
         Left            =   1200
         List            =   "frmFuel.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   670
         Width           =   1700
      End
      Begin VB.CommandButton cmdFuelCalc 
         Caption         =   "Allocate Fuel between Tanks"
         Height          =   255
         Left            =   1680
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3450
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtReserveFuel 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   23
         Text            =   "0"
         Top             =   3060
         Width           =   735
      End
      Begin VB.TextBox txtWindSpeed 
         Height          =   285
         Left            =   4080
         TabIndex        =   22
         Text            =   "0"
         Top             =   2715
         Width           =   735
      End
      Begin VB.TextBox txtCruiseSpeed 
         Height          =   285
         Left            =   1200
         TabIndex        =   21
         Top             =   2715
         Width           =   735
      End
      Begin VB.TextBox txtTaxiFuel 
         Height          =   285
         Left            =   4080
         TabIndex        =   20
         Top             =   2370
         Width           =   735
      End
      Begin VB.TextBox txtBaseFuel 
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   2370
         Width           =   735
      End
      Begin VB.TextBox txtEngineType 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1455
      End
      Begin VB.TextBox txtEquipment 
         Enabled         =   0   'False
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
         Left            =   1200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1575
      End
      Begin VB.TextBox txtAirportInfo 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Airport Name (XXXX) - Aiport Name (XXXX)"
         Top             =   325
         Width           =   4215
      End
      Begin VB.Label Label17 
         Caption         =   "pounds"
         Height          =   255
         Left            =   5280
         TabIndex        =   64
         Top             =   3090
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "FUEL REQUIRED"
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
         Left            =   2640
         TabIndex        =   62
         Top             =   3090
         Width           =   1335
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Profile Name"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1320
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000012&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   6720
         Y1              =   1120
         Y2              =   1120
      End
      Begin VB.Label Label19 
         Caption         =   "pounds"
         Height          =   255
         Left            =   5160
         TabIndex        =   55
         Top             =   2055
         Width           =   735
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Max Weight"
         Height          =   255
         Left            =   3120
         TabIndex        =   53
         Top             =   2055
         Width           =   855
      End
      Begin VB.Label lblEngineCount 
         Caption         =   "x 4"
         Height          =   255
         Left            =   5640
         TabIndex        =   52
         Top             =   1690
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "lbs/hr/engine"
         Height          =   255
         Left            =   2040
         TabIndex        =   51
         Top             =   2055
         Width           =   975
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Fuel Flow"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2055
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "pounds"
         Height          =   255
         Left            =   2040
         TabIndex        =   28
         Top             =   3090
         Width           =   615
      End
      Begin VB.Label Label15 
         Caption         =   "knots"
         Height          =   255
         Left            =   4920
         TabIndex        =   27
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "pounds"
         Height          =   255
         Left            =   4920
         TabIndex        =   26
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "knots"
         Height          =   255
         Left            =   2040
         TabIndex        =   25
         Top             =   2745
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "pounds"
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Reserve Fuel"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   3090
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Head Wind"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   2745
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Cruise Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2745
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxi Fuel"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Base Fuel"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lblProfile 
         Alignment       =   1  'Right Justify
         Caption         =   "Fuel Profile"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Engine Type"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Aircraft Type"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1695
         Width           =   975
      End
      Begin VB.Label lblDistance 
         Alignment       =   2  'Center
         Caption         =   "XX,XXX Miles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Flight Route"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmFuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sliders As Variant
Private SliderLabels As Variant

Private CurrentProfile As AircraftFuel

Private Sub MGWCheck()
    Dim x As Integer
    Dim TotalLoad As Long
    
    'Calculate total fuel load
    For x = 0 To UBound(Sliders)
        If Sliders(x).Enabled Then TotalLoad = TotalLoad + (Sliders(x).Max - Sliders(x).value)
    Next
    
    'Display warning
    txtTotalFuel.Text = CStr(TotalLoad)
    lblMaxWeightWarning.visible = ((TotalLoad + acInfo.ZeroFuelWeight) > acInfo.MaxGrossWeight)
End Sub

Private Sub CalcFuel()
    Dim totalFuel As Long
    Dim Distance As Integer

    'Calculate the fuel
    Distance = info.airportD.DistanceTo(info.AirportA.Latitude, info.AirportA.Longitude)
    totalFuel = CurrentProfile.CruiseFuel(Distance, CInt(txtWindSpeed.Text)) + CLng(txtReserveFuel.Text)

    'Set fuel
    txtFuelRequired.Text = CStr(totalFuel)
End Sub

Private Sub cboProfile_Click()
    Dim x As Integer

    'Lock the fields if not selected
    If (cboProfile.ListIndex = 0) Then
        LockAircraftInfo True
        frmLoader.visible = False
        Exit Sub
    End If
    
    'Populate the fields
    LoadProfile cboProfile.List(cboProfile.ListIndex)
    With CurrentProfile
        txtProfileName.Text = .Name
        txtEquipment.Text = .AircraftType
        txtEngineType.Text = .EngineType
        lblEngineCount.Caption = "x " + CStr(.EngineCount)
        txtFuelFlow.Text = CStr(.FuelFlow)
        txtBaseFuel.Text = CStr(.BaseFuel)
        txtTaxiFuel.Text = CStr(.TaxiFuel)
        txtMaxWeight.Text = CStr(acInfo.MaxGrossWeight)
        txtCruiseSpeed.Text = CStr(.CruiseSpeed)
        txtReserveFuel.Text = CStr(.FuelFlow * .EngineCount * 0.75)
        If (txtWindSpeed.Text = "") Then txtWindSpeed.Text = "0"
    End With
    
    'Reset the tank settings
    For x = 0 To (MAX_TANK - 2)
        If Sliders(x).Enabled Then Sliders(x).value = Sliders(x).Max
    Next
    
    'Make things visible
    LockAircraftInfo False
    frmLoader.visible = True
    Call CalcFuel
    
    'Reset total fuel data
    If (txtTotalFuel.Text <> "") Then
        txtTotalFuel.Text = ""
        txtTotalFuel.visible = False
        lblTotalFuel.visible = False
        lblTotalFuelPounds.visible = False
        cmdFuelLoad.visible = False
    End If
End Sub

Private Sub cmdFuelCalc_Click()
    Dim totalFuel As Long
    
    'Set critical error handler
    On Error GoTo FatalError
    
    'Load the primary/secondary tanks
    totalFuel = CLng(txtFuelRequired.Text)
    With CurrentProfile
        totalFuel = SetTanks(totalFuel, .PrimaryTanks, .PrimaryPercentage)
        If (totalFuel > 0) Then totalFuel = SetTanks(totalFuel, .SecondaryTanks, .SecondaryPercentage)
        If (totalFuel > 0) Then totalFuel = SetTanks(totalFuel, .OtherTanks, 100)
    End With
    
    'Check if we still have fuel left
    If (totalFuel > 100) Then MsgBox "Fuel Calculation is complete, but not all required fuel has been" & _
        vbCrLf & "loaded onto your Aircraft! You may need to add fuel manually.", vbExclamation, _
            "Fuel Remaining"
            
    'Make entries visible
    Call MGWCheck
    lblTotalFuel.visible = True
    txtTotalFuel.visible = True
    lblTotalFuelPounds.visible = True
    lblMaxFuel.visible = True
    txtMaxFuel.visible = True
    lblMaxFuelPounds.visible = True
    cmdFuelLoad.visible = True
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Error Calculating Fuel Load!", vbCritical, "Fuel Error"
    Resume ExitSub
End Sub

Private Function SetTanks(ByVal Fuel As Long, tNames As Variant, ByVal maxPct As Integer) As Long
    Dim tankName As Variant
    Dim Capacity As Long, TankLoad As Long
    Dim ratio As Double
    Dim tCode As Integer
    
    'Do nothing if no tanks specified
    If (IsEmpty(tNames) Or (UBound(tNames) = -1)) Then
        SetTanks = Fuel
        Exit Function
    End If
    
    'Determine total capacity
    For Each tankName In tNames
        tCode = GetTankCode(tankName)
        Capacity = Capacity + acInfo.TankCapacity(tCode)
    Next
    
    'Determine what percentage of these tanks should be filled
    ratio = Fuel * 100# / Capacity
    If (ratio > CurrentProfile.PrimaryPercentage) Then ratio = CurrentProfile.PrimaryPercentage

    'Set the tank amounts correspondingly
    For Each tankName In tNames
        tCode = GetTankCode(tankName)
        If (tCode <> -1) Then
            TankLoad = CLng(acInfo.TankCapacity(tCode) * ratio / 100)
            Sliders(tCode).value = Sliders(tCode).Max - TankLoad
            Fuel = Fuel - TankLoad
        End If
    Next

    'Return fuel load remaining
    SetTanks = Fuel
End Function

Private Sub cmdFuelLoad_Click()
    Dim x As Integer
    Dim dwResult As Long
    Dim TankLoad As Long
    Dim TankPct(10) As Long
    Dim ProfileName As String
    
    'Set critical error handler
    On Error GoTo FatalError
    
    'Load the fuel tanks
    frmFuel.MousePointer = vbHourglass
    For x = 0 To (MAX_TANK - 2)
        If Sliders(x).Enabled Then
            TankLoad = (Sliders(x).Max - Sliders(x).value)
            If config.ShowDebug Then ShowMessage "Loading " & TankNames(x) & " with " & CStr(TankLoad) & _
                " lbs Fuel", DEBUGTEXTCOLOR
                
            'Write fuel tank level as (% * 128 * 65536) of capacity
            TankPct(x) = TankLoad * 65536 / acInfo.TankCapacity(x) * 128
            Call FSUIPC_Write(CLng(AllTankOffsets(x)), 4, VarPtr(TankPct(x)), dwResult)
        End If
    Next
    
    'Make the FSUIPC call
    If Not FSUIPC_Process(dwResult) Then ShowMessage "Error loading Fuel Tanks!", ACARSERRORCOLOR
    
    'Update the aircraft's fuel profile
    ProfileName = cboProfile.List(cboProfile.ListIndex)
    If (acInfo.FuelProfile <> ProfileName) Then
        If (MsgBox("Do you want to update this aircraft's default Fuel Loading profile?", _
            vbQuestion Or vbYesNo, "Update Fuel Profile") = vbYes) Then WriteINI "General", _
            "ACARSFuelProfile", ProfileName, acInfo.CFGFile
    End If
    
    info.FuelLoaded = True
    MsgBox "Fuel Loaded onto Aircraft. ", vbInformation, "Fuel Loaded"
    Unload frmFuel
    
ExitSub:
    frmFuel.MousePointer = vbDefault
    Exit Sub
    
FatalError:
    ShowMessage "Error Loading " & TankNames(x) & " - " + err.Description, ACARSERRORCOLOR
    Resume ExitSub
End Sub

Private Sub Form_Load()
    Dim x As Integer
    Dim Distance As Integer
    Dim MaxFuel As Long
    
    'Load the airports and distance
    Distance = info.airportD.DistanceTo(info.AirportA.Latitude, info.AirportA.Longitude)
    txtAirportInfo.Text = info.airportD.Name + " (" + info.airportD.ICAO + ") - " + _
        info.AirportA.Name + " (" + info.AirportA.ICAO + ")"
    lblDistance.Caption = CStr(Distance) + " Miles"
    
    'Load slider arrays
    Sliders = Array(sldCenter, sldLeftMain, sldLeftAUX, sldLeftTip, sldRightMain, sldRightAUX, _
        sldRightTip, sldCenter2, sldCenter3)
    SliderLabels = Array(lblCenter, lblLeftMain, lblLeftAUX, lblLeftTip, lblRightMain, lblRightAUX, _
        lblRightTip, lblCenter2, lblCenter3)

    'Load aircraft information
    If (acInfo Is Nothing) Then Set acInfo = GetAircraftInfo()
    For x = 0 To MAX_TANK
        MaxFuel = MaxFuel + acInfo.TankCapacity(x)
    Next
        
    'Selectively enable tanks - no external tank support
    For x = 0 To (MAX_TANK - 2)
        If acInfo.HasTank(x) Then
            Sliders(x).visible = True
            Sliders(x).Enabled = True
            SliderLabels(x).Enabled = True
            Sliders(x).Max = acInfo.TankCapacity(x)
        Else
            Sliders(x).Enabled = False
            Sliders(x).visible = False
            SliderLabels(x).Enabled = False
            Sliders(x).Max = 10
        End If
        
        Sliders(x).TickFrequency = (Sliders(x).Max \ 10)
    Next
    
    'Load the profile sections
    SetComboChoices cboProfile, GetINISections(App.path + "\" + FUEL_PROFILE_FILE), _
        acInfo.FuelProfile, "< SELECT >"
    LockAircraftInfo True
    If (cboProfile.ListIndex > 0) Then Call cboProfile_Click
    txtMaxFuel.Text = CStr(MaxFuel)

End Sub

Private Sub LoadProfile(ByVal ProfileName As String)
    Dim pFile As String
    Dim profile As New AircraftFuel
    
    'Set critical error handler
    On Error GoTo FatalError
    
    'Populate fields
    pFile = App.path + "\" + FUEL_PROFILE_FILE
    With profile
        .Name = ReadINI(ProfileName, "Name", "", pFile)
        .AircraftType = ReadINI(ProfileName, "Aircraft", "", pFile)
        .EngineCount = CInt(ReadINI(ProfileName, "Engines", "2", pFile))
        .EngineType = ReadINI(ProfileName, "EngineType", "?", pFile)
        .CruiseSpeed = CInt(ReadINI(ProfileName, "CruiseSpeed", "350", pFile))
        .FuelFlow = CLng(ReadINI(ProfileName, "FuelFlow", "10000", pFile))
        .BaseFuel = CLng(ReadINI(ProfileName, "BaseFuel", "0", pFile))
        .TaxiFuel = CLng(ReadINI(ProfileName, "TaxiFuel", "1000", pFile))
        
        'Load tank data
        .PrimaryTanks = Split(ReadINI(ProfileName, "PrimaryTanks", "Left Main,Right Main", pFile), ",")
        .PrimaryPercentage = CInt(ReadINI(ProfileName, "PrimaryPercentage", "100", pFile))
        .SecondaryTanks = Split(ReadINI(ProfileName, "SecondaryTanks", "", pFile), ",")
        .SecondaryPercentage = CInt(ReadINI(ProfileName, "SecondaryPercentage", "0", pFile))
        .OtherTanks = Split(ReadINI(ProfileName, "OtherTanks", "", pFile), ",")
    End With
    
    'Save the data
    Set CurrentProfile = profile
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Unable to load " & ProfileName & " fuel Profile", vbCritical, "I/O Error"
    Resume ExitSub
        
End Sub

Private Sub LockAircraftInfo(Optional ByVal LockIt As Boolean = True)
    txtBaseFuel.Enabled = Not LockIt
    txtTaxiFuel.Enabled = Not LockIt
    txtCruiseSpeed.Enabled = Not LockIt
    txtWindSpeed.Enabled = Not LockIt
    txtReserveFuel.Enabled = Not LockIt
    cmdFuelCalc.visible = Not LockIt
    lblEngineCount.visible = Not LockIt
End Sub

Private Sub sldLeftTip_Scroll()
    sldLeftTip.Text = CStr(sldLeftTip.Max - sldLeftTip.value)
    Call MGWCheck
End Sub

Private Sub sldLeftAUX_Scroll()
    sldLeftAUX.Text = CStr(sldLeftAUX.Max - sldLeftAUX.value)
    Call MGWCheck
End Sub

Private Sub sldLeftMain_Scroll()
    sldLeftMain.Text = CStr(sldLeftMain.Max - sldLeftMain.value)
    Call MGWCheck
End Sub

Private Sub sldCenter_Scroll()
    sldCenter.Text = CStr(sldCenter.Max - sldCenter.value)
    Call MGWCheck
End Sub

Private Sub sldCenter2_Scroll()
    sldCenter2.Text = CStr(sldCenter2.Max - sldCenter2.value)
    Call MGWCheck
End Sub

Private Sub sldCenter3_Scroll()
    sldCenter3.Text = CStr(sldCenter3.Max - sldCenter3.value)
    Call MGWCheck
End Sub

Private Sub sldRightMain_Scroll()
    sldRightMain.Text = CStr(sldRightMain.Max - sldRightMain.value)
    Call MGWCheck
End Sub

Private Sub sldRightAUX_Scroll()
    sldRightAUX.Text = CStr(sldRightAUX.Max - sldRightAUX.value)
    Call MGWCheck
End Sub

Private Sub sldRightTip_Scroll()
    sldRightTip.Text = CStr(sldRightTip.Max - sldRightTip.value)
    Call MGWCheck
End Sub

Private Sub txtBaseFuel_Change()
    LimitNumber txtBaseFuel, 25000
    CurrentProfile.BaseFuel = CLng(txtBaseFuel.Text)
    Call CalcFuel
End Sub

Private Sub txtCruiseSpeed_Change()
    LimitNumber txtCruiseSpeed, 1350
    CurrentProfile.CruiseSpeed = CInt(txtCruiseSpeed.Text)
    Call CalcFuel
End Sub

Private Sub txtReserveFuel_Change()
    LimitNumber txtReserveFuel, 40000
    Call CalcFuel
End Sub

Private Sub txtTaxiFuel_Change()
    LimitNumber txtTaxiFuel, 9000
    CurrentProfile.TaxiFuel = CLng(txtTaxiFuel.Text)
    Call CalcFuel
End Sub

Private Sub txtWindSpeed_Change()
    LimitNumber txtWindSpeed, 250
    Call CalcFuel
End Sub

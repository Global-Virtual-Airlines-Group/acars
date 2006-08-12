VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFailure 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Failure Configuration"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFailure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sldFailureCount 
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Frame frmEquipment 
      Caption         =   "Aircraft Equipment Failures"
      Height          =   1950
      Left            =   120
      TabIndex        =   1
      Top             =   3030
      Width           =   5655
      Begin VB.CheckBox chkReverser4 
         Caption         =   "Thrust Reverser #4"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   27
         Top             =   1620
         Width           =   1800
      End
      Begin VB.CheckBox chkReverser3 
         Caption         =   "Thrust Reverser #3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   1360
         Width           =   1800
      End
      Begin VB.CheckBox chkReverser2 
         Caption         =   "Thrust Reverser #2"
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   1100
         Width           =   1800
      End
      Begin VB.CheckBox chkReverser1 
         Caption         =   "Thrust Reverser #1"
         Height          =   255
         Left            =   3730
         TabIndex        =   24
         Top             =   840
         Width           =   1800
      End
      Begin VB.CheckBox chkPitot 
         Caption         =   "Pitot Heat"
         Height          =   255
         Left            =   2100
         TabIndex        =   23
         Top             =   1620
         Width           =   1455
      End
      Begin VB.CheckBox chkGear 
         Caption         =   "Landing Gear"
         Height          =   255
         Left            =   2100
         TabIndex        =   22
         Top             =   1360
         Width           =   1455
      End
      Begin VB.CheckBox chkSpoilers 
         Caption         =   "Spoilers"
         Height          =   255
         Left            =   2100
         TabIndex        =   21
         Top             =   1100
         Width           =   1455
      End
      Begin VB.CheckBox chkFlaps 
         Caption         =   "Wing Flaps"
         Height          =   255
         Left            =   2100
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkEngine4 
         Caption         =   "Engine #4"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1620
         Width           =   1455
      End
      Begin VB.CheckBox chkEngine3 
         Caption         =   "Engine #3"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1360
         Width           =   1455
      End
      Begin VB.CheckBox chkEngine2 
         Caption         =   "Engine #2"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1100
         Width           =   1455
      End
      Begin VB.CheckBox chkEngine1 
         Caption         =   "Engine #1"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldEquipFailure 
         Height          =   255
         Left            =   2760
         TabIndex        =   15
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickFrequency   =   10
         TextPosition    =   1
      End
      Begin VB.Label lblCompFailureTypes 
         Caption         =   "Allow the following Equipment failures:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lblEQProbability 
         Caption         =   "Failure Probability"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   270
         Width           =   2655
      End
   End
   Begin VB.Frame frmInst 
      Caption         =   "Instrumentation Failures"
      Height          =   1950
      Left            =   120
      TabIndex        =   28
      Top             =   960
      Width           =   5655
      Begin VB.CheckBox chkTransponder 
         Caption         =   "Transponder"
         Height          =   255
         Left            =   3730
         TabIndex        =   13
         Top             =   1360
         Width           =   1575
      End
      Begin VB.CheckBox chkASI 
         Caption         =   "Airspeed Indicator"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1620
         Width           =   1695
      End
      Begin VB.CheckBox chkVSI 
         Caption         =   "Vertical Speed Gauge"
         Height          =   255
         Left            =   3730
         TabIndex        =   14
         Top             =   1620
         Width           =   1890
      End
      Begin VB.CheckBox chkNAV2 
         Caption         =   "NAV2 Radio"
         Height          =   255
         Left            =   3730
         TabIndex        =   12
         Top             =   1100
         Width           =   1575
      End
      Begin VB.CheckBox chkNAV1 
         Caption         =   "NAV1 Radio"
         Height          =   255
         Left            =   3730
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkFuel 
         Caption         =   "Fuel Gauge"
         Height          =   255
         Left            =   2100
         TabIndex        =   10
         Top             =   1620
         Width           =   1575
      End
      Begin VB.CheckBox chkCompass 
         Caption         =   "Compass"
         Height          =   255
         Left            =   2100
         TabIndex        =   9
         Top             =   1360
         Width           =   1575
      End
      Begin VB.CheckBox chkCOM2 
         Caption         =   "COM2 Radio"
         Height          =   255
         Left            =   2100
         TabIndex        =   8
         Top             =   1100
         Width           =   1455
      End
      Begin VB.CheckBox chkCOM1 
         Caption         =   "COM1 Radio"
         Height          =   255
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkAttitude 
         Caption         =   "Attitude Indicator"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1360
         Width           =   1695
      End
      Begin VB.CheckBox chkAltimeter 
         Caption         =   "Altimeter"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1100
         Width           =   1455
      End
      Begin VB.CheckBox chkADF 
         Caption         =   "ADF Radio"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin MSComctlLib.Slider sldInstFailure 
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickFrequency   =   10
         TextPosition    =   1
      End
      Begin VB.Label lblInstFailureTypes 
         Caption         =   "Allow the following Instrument failures:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label lblInstProbability 
         Caption         =   "Failure Probability"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   270
         Width           =   2655
      End
   End
   Begin VB.Label lblMaxFailCount 
      Caption         =   "0"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   600
      Width           =   255
   End
   Begin VB.Label lblMaxFailures 
      Caption         =   "Maximum Number of Failures per Flight"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Select the type of equipment failures you wish to simulate at random. Failures may only be simulated on Training Flights."
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmFailure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim fInfo As Failures
    Dim EngineCount As Integer
        
    'Determine the number of engines
    EngineCount = GetEngineCount(info.EquipmentType)
    
    'Get failure data
    Set fInfo = config.FailureConfig
    
    'Set sliders
    sldFailureCount.value = fInfo.MaxFailures
    sldInstFailure.value = fInfo.InstrumentFailureProbability
    sldEquipFailure.value = fInfo.EquipmentFailureProbability
    
    'Set instrumentation check box fields
    If fInfo.ADF Then chkADF.value = 1
    If fInfo.Altimeter Then chkAltimeter.value = 1
    If fInfo.Attitude Then chkAttitude.value = 1
    If fInfo.Airspeed Then chkASI.value = 1
    If fInfo.COM1 Then chkCOM1.value = 1
    If fInfo.COM2 Then chkCOM2.value = 1
    If fInfo.Compass Then chkCompass.value = 1
    If fInfo.Fuel Then chkFuel.value = 1
    If fInfo.NAV1 Then chkNAV1.value = 1
    If fInfo.NAV2 Then chkNAV2.value = 1
    If fInfo.Transponder Then chkTransponder.value = 1
    If fInfo.VerticalSpeed Then chkVSI.value = 1
    
    'Set equipment check box fields
    If fInfo.Flaps Then chkFlaps.value = 1
    If fInfo.Spoilers Then chkSpoilers.value = 1
    If fInfo.Gear Then chkGear.value = 1
    If fInfo.PitotHeat Then chkPitot.value = 1
    If fInfo.EngineFailure(1) Then chkEngine1.value = 1
    If fInfo.ReverserFailure(1) Then chkReverser1.value = 1
    If fInfo.EngineFailure(2) Then chkEngine2.value = 1
    If fInfo.ReverserFailure(2) Then chkReverser2.value = 1
    If (EngineCount = 3) Then
        If fInfo.EngineFailure(3) Then chkEngine3.value = 1
        If fInfo.ReverserFailure(3) Then chkReverser3.value = 1
        chkEngine3.Enabled = True
        chkReverser3.Enabled = True
        chkEngine4.Enabled = False
        chkReverser4.Enabled = False
    ElseIf (EngineCount = 4) Then
        If fInfo.EngineFailure(3) Then chkEngine3.value = 1
        If fInfo.ReverserFailure(3) Then chkReverser3.value = 1
        If fInfo.EngineFailure(4) Then chkEngine4.value = 1
        If fInfo.ReverserFailure(4) Then chkReverser4.value = 1
        chkEngine3.Enabled = True
        chkReverser3.Enabled = True
        chkEngine4.Enabled = True
        chkReverser4.Enabled = True
    Else
        chkEngine3.Enabled = False
        chkReverser3.Enabled = False
        chkEngine4.Enabled = False
        chkReverser4.Enabled = False
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim fInfo As New Failures

    fInfo.InstrumentFailureProbability = sldInstFailure.value
    fInfo.EquipmentFailureProbability = sldEquipFailure.value
    fInfo.MaxFailures = sldFailureCount.value
        
    'Load instrumentation checkboxes
    fInfo.ADF = (chkADF.value = 1)
    fInfo.Altimeter = (chkAltimeter.value = 1)
    fInfo.Attitude = (chkAttitude.value = 1)
    fInfo.Airspeed = (chkASI.value = 1)
    fInfo.COM1 = (chkCOM1.value = 1)
    fInfo.COM2 = (chkCOM2.value = 1)
    fInfo.Compass = (chkCompass.value = 1)
    fInfo.Fuel = (chkFuel.value = 1)
    fInfo.NAV1 = (chkNAV1.value = 1)
    fInfo.NAV2 = (chkNAV2.value = 1)
    fInfo.Transponder = (chkTransponder.value = 1)
    fInfo.VerticalSpeed = (chkVSI.value = 1)
        
    'Load equipment checkboxes
    fInfo.Flaps = (chkFlaps.value = 1)
    fInfo.Spoilers = (chkFlaps.value = 1)
    fInfo.Gear = (chkGear.value = 1)
    fInfo.PitotHeat = (chkPitot.value = 1)
    fInfo.SetEngineFailure 1, (chkEngine1.value = 1)
    fInfo.SetReverserFailure 1, (chkReverser1.value = 1)
    fInfo.SetEngineFailure 2, (chkEngine2.value = 1)
    fInfo.SetReverserFailure 2, (chkReverser2.value = 1)
    fInfo.SetEngineFailure 3, (chkEngine3.value = 1)
    fInfo.SetReverserFailure 3, (chkReverser3.value = 1)
    fInfo.SetEngineFailure 4, (chkEngine4.value = 1)
    fInfo.SetReverserFailure 4, (chkReverser4.value = 1)
        
    'Save configuration
    Set config.FailureConfig = fInfo
End Sub

Private Function GetEngineCount(ByVal eqType As String) As Integer
    Dim x As Integer
    Dim Eng3Aircraft As Variant
    Dim Eng4Aircraft As Variant

    'Load aircraft types
    Eng3Aircraft = Array("B727-100", "B727-200", "L-1011", "DC-10-10", "DC-10-30", "DC-10-40", "MD-11")
    Eng4Aircraft = Array("A340-300", "A340-600", "B707-120", "B707-320", "B720", "B747-100", _
        "B747-200", "B747-300", "B747-400", "BAE-146", "CV-880", "CV-990", "Comet", "Constellation", _
        "DC-6", "DC-7", "DC-8-11", "DC-8-21", "DC-8-33", "DC-8-42", "DC-8-51", "DC-8-61", _
        "DC-8-62", "DC-8-71", "DC-8-72", "L-100", "-")

    'Check if we are a 3-holer
    For x = 0 To UBound(Eng3Aircraft)
        If (Eng3Aircraft(x) = eqType) Then
            GetEngineCount = 3
            Exit Function
        End If
    Next

    'Check if we are a 4-holer
    For x = 0 To UBound(Eng4Aircraft)
        If (Eng4Aircraft(x) = eqType) Then
            GetEngineCount = 4
            Exit Function
        End If
    Next
    
    GetEngineCount = 2
End Function

Private Sub sldEquipFailure_Change()
    lblEQProbability.Caption = "Failure Probability (" + CStr(sldEquipFailure.value) + "% per hour)"
End Sub

Private Sub sldFailureCount_Change()
    lblMaxFailCount.Caption = CStr(sldFailureCount.value)
End Sub

Private Sub sldInstFailure_Change()
    lblInstProbability.Caption = "Failure Probability (" + CStr(sldInstFailure.value) + "% per hour)"
End Sub


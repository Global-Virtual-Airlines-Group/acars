VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DVA ACARS"
   ClientHeight    =   7125
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9435
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
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar PositionProgress 
      Height          =   255
      Left            =   6720
      TabIndex        =   30
      Top             =   1410
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.ComboBox cboNetwork 
      Height          =   315
      ItemData        =   "frmMain.frx":0442
      Left            =   6720
      List            =   "frmMain.frx":0452
      TabIndex        =   29
      Text            =   "Offline"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboAirportA 
      Height          =   315
      ItemData        =   "frmMain.frx":0472
      Left            =   1440
      List            =   "frmMain.frx":0479
      TabIndex        =   9
      Text            =   "cboAirportA"
      Top             =   1320
      Width           =   3975
   End
   Begin VB.ComboBox cboAirportD 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":0489
      Left            =   1440
      List            =   "frmMain.frx":0490
      TabIndex        =   8
      Text            =   "cboAirportD"
      Top             =   960
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Tahoma"
   End
   Begin VB.Timer tmrFlightTimeCounter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1470
      Top             =   3600
   End
   Begin VB.Timer tmrRequest 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1035
      Top             =   3600
   End
   Begin VB.Timer tmrPosUpdates 
      Interval        =   250
      Left            =   600
      Top             =   3600
   End
   Begin MSWinsockLib.Winsock wsckMain 
      Left            =   165
      Top             =   3600
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
      Left            =   8265
      TabIndex        =   12
      Top             =   45
      Width           =   1095
   End
   Begin VB.ListBox lstPilots 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3060
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":04A0
      Left            =   7935
      List            =   "frmMain.frx":04A2
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   1440
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   285
      Left            =   8460
      TabIndex        =   1
      Top             =   6480
      Width           =   915
   End
   Begin VB.TextBox txtCmd 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   6480
      Width           =   8340
   End
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   3075
      Left            =   60
      TabIndex        =   15
      Top             =   3360
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   5424
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":04A4
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   27
      Top             =   6825
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5822
            MinWidth        =   5822
            Text            =   "Status: Connected to ACARS server"
            TextSave        =   "Status: Connected to ACARS server"
            Key             =   "status"
            Object.Tag             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3351
            MinWidth        =   3351
            Text            =   "Flight Phase: Airborne"
            TextSave        =   "Flight Phase: Airborne"
            Key             =   "flightphase"
            Object.Tag             =   "flightphase"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4322
            MinWidth        =   4322
            Text            =   "Last Position Report: 02:38:08 Z"
            TextSave        =   "Last Position Report: 02:38:08 Z"
            Key             =   "lastposrep"
            Object.Tag             =   "lastposrep"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Sim Time: 02:38:26 Z"
            TextSave        =   "Sim Time: 02:38:26 Z"
            Key             =   "simtime"
            Object.Tag             =   "simtime"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboEquipment 
      Height          =   315
      ItemData        =   "frmMain.frx":051F
      Left            =   3960
      List            =   "frmMain.frx":0521
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdPIREP 
      Caption         =   "File PIREP"
      Enabled         =   0   'False
      Height          =   345
      Left            =   8265
      TabIndex        =   14
      Top             =   795
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdStartStopFlight 
      Caption         =   "Start Flight"
      Height          =   345
      Left            =   8265
      TabIndex        =   13
      Top             =   420
      Width           =   1095
   End
   Begin VB.TextBox txtCruiseAlt 
      Height          =   285
      Left            =   6720
      TabIndex        =   7
      Top             =   480
      Width           =   1000
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
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   120
      Width           =   1200
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
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFlightNumber 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   675
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2520
      Width           =   7935
   End
   Begin VB.TextBox txtRoute 
      Height          =   675
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1800
      Width           =   7935
   End
   Begin VB.TextBox txtPilotName 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label progressLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "OFFLINE PIREP"
      Height          =   255
      Left            =   5400
      TabIndex        =   31
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Network:"
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      Top             =   1010
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Arriving at:"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Departing from:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1000
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   255
      Left            =   5745
      TabIndex        =   26
      Top             =   165
      Width           =   915
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Cruise Altitude:"
      Height          =   255
      Left            =   5445
      TabIndex        =   25
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Pilot ID:"
      Height          =   255
      Left            =   3285
      TabIndex        =   24
      Top             =   165
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipment:"
      Height          =   255
      Left            =   3045
      TabIndex        =   23
      Top             =   525
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Remarks:"
      Height          =   255
      Left            =   285
      TabIndex        =   21
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Flight Route:"
      Height          =   255
      Left            =   45
      TabIndex        =   20
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Flight Number:"
      Height          =   255
      Left            =   285
      TabIndex        =   18
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Pilot Name:"
      Height          =   255
      Left            =   285
      TabIndex        =   17
      Top             =   165
      Width           =   1095
   End
   Begin VB.Menu mnuFlight 
      Caption         =   "&Flight"
      Begin VB.Menu mnuConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuFlightStartFlight 
         Caption         =   "Start &Flight"
      End
      Begin VB.Menu mnuFlightEndFlight 
         Caption         =   "&End Flight"
      End
      Begin VB.Menu mnuFlightSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenFlightPlan 
         Caption         =   "&Open Flight Plan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveFlightPlan 
         Caption         =   "&Save SB3 Flight Plan"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFlightSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlightExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsPlaySoundForChat 
         Caption         =   "Play &Sound For Chat"
      End
      Begin VB.Menu mnuOptionsFlyOffline 
         Caption         =   "Fly &Offline"
      End
      Begin VB.Menu mnuOptionsSavePassword 
         Caption         =   "Save &Password"
      End
      Begin VB.Menu mnuOptionsShowTimestamps 
         Caption         =   "Show &Timestamps"
      End
      Begin VB.Menu mnuOptionsShowDebugMessages 
         Caption         =   "Show &Debug Messages"
      End
      Begin VB.Menu mnuOptionsSB3Support 
         Caption         =   "SquawkBox 3 Integration"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngWindowHandle As Long

Dim blnACARSConnectionResolved As Boolean
Dim CmdHistory() As String
Dim intCmdHistIndex As Integer
Dim blnFirstCmd As Boolean
Dim strCmdBuffer As String
Dim intCmdBufferSel As Integer
Dim blnUserQuit As Boolean

Private Sub cboAirportA_Click()
    If (cboAirportA.ListIndex > 0) Then info.AirportA = config.AirportCodes(cboAirportA.ListIndex - 1)
End Sub

Private Sub cboAirportD_Click()
    If (cboAirportD.ListIndex > 0) Then info.AirportD = config.AirportCodes(cboAirportD.ListIndex - 1)
End Sub

Private Sub cboEquipment_Click()
    info.EquipmentType = cboEquipment.List(cboEquipment.ListIndex)
End Sub

Private Sub cboNetwork_Change()
    info.Network = cboNetwork.List(cboNetwork.ListIndex)
End Sub

Private Sub cmdConnectDisconnect_Click()
    ToggleACARSConnection
End Sub

Private Sub cmdPIREP_Click()

    'If we're not connected, then stop
    If Not config.ACARSConnected Then
        MsgBox "You must be connected to the ACARS server to file a Flight Report.", vbExclamation + vbOKOnly, "Not Connected"
        frmMain.cmdPIREP.Visible = False
        Exit Sub
    End If
    
    'Make sure that there is data
    If Not info.FlightData Then
        MsgBox "The Flight must be completed before you can file a Flight Report.", vbExclamation + vbOKOnly, "Incomplete Flight"
        Exit Sub
    End If
    
    'If we've already filed the PIREP and it was approved, then stop
    If info.PIREPFiled Then Exit Sub
    
    'If we're offline, then file the Flight Info
    frmMain.cmdPIREP.Enabled = False
    If info.Offline Then
        Dim Queue As Variant
        Dim x As Integer
    
        If Not positions.HasData Then
            MsgBox "No Offline Flight data Available.", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        
        'Display the progress bar
        Queue = positions.Queue
        frmMain.progressLabel.Visible = True
        frmMain.PositionProgress.value = 0
        frmMain.PositionProgress.Max = UBound(Queue) + 2
        frmMain.PositionProgress.Visible = True
    
        'Send the flight data
        If (info.flightID = 0) Then
            SendFlightInfo info
            ReqStack.Send
            ShowMessage "Sending Offline Flight Information", ACARSTEXTCOLOR
        
            'Wait for the flight ID to be returned
            Dim totalSleep As Integer
            While ((info.flightID = 0) And (totalSleep < 6000))
                totalSleep = totalSleep + 250
                Sleep 250
                DoEvents
            Wend
        
            'If we have waited too long, abort
            If (info.flightID = 0) Then
                MsgBox "ACARS Server timed out returning Flight ID", vbCritical + vbOKOnly
                frmMain.progressLabel.Visible = True
                frmMain.PositionProgress.Visible = True
                frmMain.cmdPIREP.Enabled = True
                Exit Sub
            End If
        End If
        
        'File the Flight Positions
        ShowMessage "Sending " + CStr(UBound(Queue) + 1) + " Position Records", ACARSTEXTCOLOR
        For x = 0 To UBound(Queue)
            SendPosition Queue(x)
            frmMain.PositionProgress.value = x
            If ((x Mod 3) = 0) Then
                ReqStack.Send
                frmMain.PositionProgress.Refresh
                Sleep 375
                DoEvents
            End If
        Next
        
        'Empty out the position data
        ReqStack.Send
        Sleep 375
        DoEvents
        
        'Mark the flight as ended and empty the request stack
        SendEndFlight
        ReqStack.Send
        Sleep 450
        DoEvents
    End If

    'Send the PIREP
    info.PIREPReqID = SendPIREP(info)
    ShowMessage "Sending Flight Report " + Hex(info.PIREPReqID), ACARSTEXTCOLOR
    ReqStack.Send
    frmMain.PositionProgress.value = frmMain.PositionProgress.Max
    DoEvents
End Sub

Private Sub cmdSend_Click()
    txtCmd_KeyPress 13
End Sub

Private Sub cmdStartStopFlight_Click()
    ToggleFlight
End Sub

Private Sub Form_Load()
    intCmdHistIndex = 0
    blnFirstCmd = True
    lngWindowHandle = Me.hWnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ConfirmExit Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Close FSUIPC connection. (No harm trying to close it when it's not actually open.)
    FSUIPC_Close
    
    'Save settings
    config.Save

    'Unload all forms.
    Dim i As Integer
    For i = (Forms.Count - 1) To 0 Step -1
        Unload Forms(i)
    Next
    
    End
End Sub

Private Sub mnuConnect_Click()
    ToggleACARSConnection
End Sub

Private Sub mnuFlightEndFlight_Click()
    ToggleFlight
End Sub

Private Sub mnuFlightExit_Click()
    Unload frmMain
End Sub

Private Sub mnuFlightStartFlight_Click()
    ToggleFlight
End Sub

Private Sub mnuOpenFlightPlan_Click()
    FPlan_Open
    config.UpdateFlightInfo
End Sub

Private Sub mnuOptionsSB3Support_Click()
    config.SB3Support = Not config.SB3Support
    config.UpdateSettingsMenu
End Sub

Private Sub mnuSaveFlightPlan_Click()
    SB3Plan_Save
End Sub

Private Sub mnuOptionsFlyOffline_Click()
    config.FlyOffline = Not config.FlyOffline
    config.UpdateSettingsMenu
End Sub

Private Sub mnuOptionsPlaySoundForChat_Click()
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

Private Sub mnuOptionsShowTimestamps_Click()
    config.ShowTimestamps = Not config.ShowTimestamps
    config.UpdateSettingsMenu
End Sub

Sub ToggleFlight()
    If info.InFlight Then StopFlight Else StartFlight
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
        intResponse = MsgBox("You have not yet completed this flight. The flight is considered complete when you have taken off, landed, taxied to your gate and set the parking brake. You won't be able to auto-file a PIREP until the flight is complete. Are you sure you want to end the flight early?", vbYesNo Or vbQuestion, "Confirm")
        If intResponse = vbNo Then
            StopFlight = False
            Exit Function
        End If
    End If

    'Stop tracking/timing flight.
    tmrPosUpdates.Enabled = False
    tmrFlightTimeCounter.Enabled = False

    'Set some flags and control states.
    If info.FlightData Then
        info.Phase = "Completed"
    Else
        If isError Then
            info.Phase = "Error"
        Else
            info.Phase = "Aborted"
        End If
    End If
    
    sbMain.Panels(2).Text = "Flight Phase: " & info.Phase
    info.InFlight = False

    'Clear the flight ID registry entry.
    config.SaveFlightID 0

    'Send end_flight request if connected.
    If config.ACARSConnected Then SendEndFlight

    'Close FSUIPC link.
    FSUIPC_Close
    config.FSUIPCConnected = False

    'Show flight info window if the flight was completed.
    SetButtonMenuStates
    If (info.Phase = "Completed") Then
        cmdPIREP.Visible = True
        cmdPIREP.Enabled = config.ACARSConnected
    End If
End Function

Sub StartFlight()

    'This should never happen:
    If info.InFlight Then
        MsgBox "You already have a flight in progress!", vbOKOnly Or vbExclamation, "Error"
        Exit Sub
    End If

    'Make sure we're not starting a flight without having PIREPped a previous flight.
    If info.FlightData And Not info.PIREPFiled Then
        If MsgBox("You have not submitted a PIREP for your previous flight. If you start a new flight now, the previous flight data will be discarded. Are you sure?", vbYesNo Or vbQuestion, "Error") = vbNo Then Exit Sub
    End If
    
    'Make sure all required flight data has been entered.
    If txtPilotID.Text = "" Then
        MsgBox "Please enter your pilot ID.", vbOKOnly Or vbExclamation, "Error"
        txtPilotID.SetFocus
        Exit Sub
    ElseIf txtPassword.Text = "" Then
        MsgBox "Please enter your password.", vbOKOnly Or vbExclamation, "Error"
        txtPassword.SetFocus
        Exit Sub
    ElseIf txtFlightNumber.Text = "" Then
        MsgBox "Please enter the flight number.", vbOKOnly Or vbExclamation, "Error"
        txtFlightNumber.SetFocus
        Exit Sub
    ElseIf cboEquipment.ListIndex = 0 Then
        MsgBox "Please select your aircraft type.", vbOKOnly Or vbExclamation, "Error"
        cboEquipment.SetFocus
        Exit Sub
    ElseIf txtCruiseAlt.Text = "" Then
        MsgBox "Please enter your cruise altitude.", vbOKOnly Or vbExclamation, "Error"
        txtCruiseAlt.SetFocus
        Exit Sub
    ElseIf cboAirportD.ListIndex = 0 Then
        MsgBox "Please enter your departure airport.", vbOKOnly Or vbExclamation, "Error"
        cboAirportD.SetFocus
        Exit Sub
    ElseIf cboAirportA.ListIndex = 0 Then
        MsgBox "Please enter your destination airport.", vbOKOnly Or vbExclamation, "Error"
        cboAirportA.SetFocus
        Exit Sub
    ElseIf txtRoute.Text = "" Then
        MsgBox "Please enter your route of flight.", vbOKOnly Or vbExclamation, "Error"
        txtRoute.SetFocus
        Exit Sub
    End If

    'Attempt to connect to FSUIPC - Make sure the FSUIPC connection succeeded.
    If Not config.FSUIPCConnected Then FSUIPCConnect
    If Not config.FSUIPCConnected Then GoTo FSERR

    'FIXME - this stuff is commented out for now since I realized I'd have to
    'store all the flight variables including phase, various timestamps, etc. every
    'time they changed, in order to be able to truly resume a flight. Future feature.
    'First check to see if the registry contains a flight ID from
    'a previous run which crashed. If so, prompt to see if the user
    'wishes to resume that flight.
    'Dim dblOldFlightID As Double
    'dblOldFlightID = CDbl(GetSetting(App.Title, "Settings", "FlightID", 0))
    'If dblOldFlightID > 0 Then
    '    If MsgBox(App.Title & " appears to have crashed during a flight. Would you like to resume that flight? Press Yes to resume the flight, No to start a new flight.", vbYesNo Or vbQuestion, "Resume Flight?") <> vbYes Then
    '        SaveSetting App.Title, "Settings", "FlightID", 0
    '        dblOldFlightID = 0
    '    End If
    'End If
    
    'Check for SB3
    If config.SB3Support Then
        config.SB3Connected = SB3Connected()
        If (SB3Running() And Not config.SB3Connected) Then
            If (MsgBox("Squawkbox 3 is running, but not connected to VATSIM. Do you want to connect?", _
                vbExclamation + vbYesNo, "Squawkbox 3") = vbYes) Then Exit Sub
        End If
        
        'If we're running, automatically set the network to VATSIM
        If config.SB3Connected Then
            frmMain.cboNetwork.ListIndex = 1
            info.Network = "VATSIM"
        End If
    End If

    'Make sure the aircraft is parked and on the ground.
    Set pos = RecordFlightData()
    If (info.flightID = 0) And ((Not pos.Parked) Or (Not pos.onGround)) Then
        MsgBox "You must be on the ground with the parking brake set in order to start a flight.", vbExclamation Or vbOKOnly, "Error"
        Exit Sub
    End If
    
    'Check if we're connected, but only if the "fly offline"
    'option is turned off. If not, then attempt to connect.
    If Not config.ACARSConnected And Not config.FlyOffline Then
        blnACARSConnectionResolved = False
        ToggleACARSConnection

        'Wait for connection results.
        Do
            DoEvents
            If blnACARSConnectionResolved Then Exit Do
        Loop While True

        'If the connection failed, prompt to fly offline.
        If Not config.ACARSConnected Then
            Dim result As VbMsgBoxResult
            result = MsgBox("The connection to the ACARS server failed. Do you wish to fly offline?", vbYesNo Or vbQuestion, "Connection Error")
            If result = vbNo Then Exit Sub
            info.Offline = True
        End If
    ElseIf config.FlyOffline Then
        info.Offline = True
    End If
    
    'Populate the flight information
    info.FlightNumber = frmMain.txtFlightNumber.Text
    info.CruiseAltitude = frmMain.txtCruiseAlt.Text

    'If we're connected to the ACARS server, send a flight info message.
    If config.ACARSConnected And Not info.Offline Then
        SendFlightInfo info
        ReqStack.Send
    End If

    'Start timing/tracking flight
    tmrFlightTimeCounter.Enabled = True
    tmrPosUpdates.Enabled = True
    
    'Start the flight
    info.StartFlight
    
    'Update status bar.
    sbMain.Panels(2).Text = "Flight Phase: " & info.Phase

    'Set some flags, variables, and control states.
    SetButtonMenuStates
    cmdPIREP.Visible = True
    cmdPIREP.Enabled = False

ExitSub:
    Exit Sub

FSERR:
    MsgBox "The following error ocurred:" & vbCrLf & vbCrLf & strFSUIPCErrorDesc, vbOKOnly Or vbCritical, "StartFlight.Error"
    If (err <> 0) Then Resume ExitSub
End Sub

'If we're airborne, keep track of time spent in time compression.
Private Sub tmrFlightTimeCounter_Timer()
    If Not (pos Is Nothing) Then Call info.UpdateFlightTime(pos.simRate / 256, 1000)
End Sub

Private Sub tmrPosUpdates_Timer()
    Static LastPosUpdate As Date

    'Get position data
    Set pos = RecordFlightData()
    If (pos Is Nothing) Then
        tmrPosUpdates.Enabled = False
        Exit Sub
    End If

    'Stop/restart the flight timer if paused or slewing
    If ((pos.Paused Or pos.Slewing) <> tmrFlightTimeCounter.Enabled) Then
        tmrFlightTimeCounter.Enabled = Not (pos.Paused Or pos.Slewing)
    End If

    'Check if the flight phase has changed
    If PhaseChanged(pos) Then sbMain.Panels(2).Text = "Flight Phase: " & info.Phase

    'If we're connected to the ACARS server, and we have been assigned
    'a flight ID for this flight by the server, check if it's time to
    'send a position update.
    If (IsEmpty(LastPosUpdate)) Or (DateDiff("s", LastPosUpdate, Now) > config.PositionInterval) Then
        LastPosUpdate = Now
        sbMain.Panels(3).Text = "Last Position Report: " & Format(LastPosUpdate, "hh:mm:ss") & " L"
        
        'Send data to the server. Otherwise save it.
        If config.ACARSConnected And (info.flightID > 0) Then
            SendPosition pos
            ReqStack.Send
        ElseIf info.Offline And IsDate(info.StartTime) Then
            positions.AddPosition pos
            If config.ShowDebug Then ShowMessage "Position Cache = " + CStr(positions.Size), DEBUGTEXTCOLOR
        End If
    End If

    'Force the sim rate to be within the allowed range.
    CheckSimRate MINTIMECOMPRESSION, MAXTIMECOMPRESSION
        
    'Calculate the update interval based on our phase/ground speed
    Dim newInterval As Integer
    newInterval = UpdatePositionInterval()
    If (config.PositionInterval <> newInterval) Then
        config.PositionInterval = newInterval
        If config.ShowDebug Then ShowMessage "Position Interval set to " + CStr(newInterval) + "s", DEBUGTEXTCOLOR
    End If
End Sub

Private Sub ToggleACARSConnection()
    If Not config.ACARSConnected Then
        cmdConnectDisconnect.Enabled = False
        On Error GoTo EH
        wsckMain.Close
        wsckMain.RemoteHost = config.ACARSHost
        wsckMain.RemotePort = config.ACARSPort
        config.SeenHELO = False
        wsckMain.Connect
        frmMain.cmdPIREP.Visible = info.FlightData
        frmMain.cmdPIREP.Enabled = info.FlightData
        frmMain.txtPilotID.Enabled = False
        frmMain.txtPassword.Enabled = False
    Else
        cmdConnectDisconnect.Enabled = False
        If Not ConfirmDisconnect Then
            cmdConnectDisconnect.Enabled = True
            Exit Sub
        End If
        CloseACARSConnection True
        cmdConnectDisconnect.Enabled = True
        frmMain.cmdPIREP.Visible = False
        frmMain.txtPilotID.Enabled = True
        frmMain.txtPassword.Enabled = True
        info.Offline = True
    End If
    
ExitSub:
    Exit Sub
    
EH:
    cmdConnectDisconnect.Enabled = True
    If ((err.Number <> 40060) And (err.Number <> 10053)) Then
        MsgBox "The following error occurred: " & err.Description & " (" & err.Number & ")", vbOKOnly Or vbCritical, "ToggleACARSConnection.Error"
    End If
    
    Resume ExitSub
End Sub

Private Sub tmrRequest_Timer()
    If Not config.ACARSConnected Then Exit Sub
    If (DateDiff("s", ReqStack.LastUse, Now) > config.PingInterval) Then
        SendPing
        ReqStack.Send
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
    info.CruiseAltitude = txtCruiseAlt.Text
End Sub

Private Sub txtPilotID_Change()
    txtPilotID.Text = UCase(txtPilotID.Text)
End Sub

Private Sub txtRemarks_Change()
    info.Remarks = txtRemarks.Text
End Sub

Private Sub txtRoute_Change()
    info.Route = UCase(txtRoute.Text)
End Sub

Private Sub wsckMain_Close()
    CloseACARSConnection False
    If Not blnUserQuit Then
        ShowMessage "ACARS connection closed by server!", ACARSERRORCOLOR
        If info.InFlight Then info.Offline = True
    End If
End Sub

Private Sub wsckMain_Connect()
    config.ACARSConnected = True
    blnACARSConnectionResolved = True
    tmrRequest.Enabled = True
    cmdConnectDisconnect.Caption = "Disconnect"
    mnuConnect.Caption = "Disconnect"
    sbMain.Panels(1).Text = "Status: Connected to ACARS server"
    
    'Log in
    info.AuthReqID = SendCredentials(frmMain.txtPilotID.Text, frmMain.txtPassword.Text)
    ReqStack.Send
    cmdConnectDisconnect.Enabled = True
End Sub

Private Sub wsckMain_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    wsckMain.GetData strData, vbString
    ProcessServerData strData
    DoEvents
End Sub

Private Sub wsckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    blnACARSConnectionResolved = True
    
    'If we get a 40006 or 10053 then toggle the ACARS connection
    If ((Number = 40006) Or (Number = 10053)) Then
        CloseACARSConnection False
        If info.InFlight Then info.Offline = True
        ShowMessage "Lost Connection to ACARS Server", ACARSERRORCOLOR
    Else
        MsgBox "The following error occurred: " & Description & " (" & Number & ")", vbOKOnly Or vbCritical, "wsckMain.Error"
        CloseACARSConnection False
    End If
End Sub

Private Function ConfirmDisconnect() As Boolean
    ConfirmDisconnect = (MsgBox("Are you sure you want to disconnect?", vbYesNo Or vbQuestion, "Confirm") = vbYes)
End Function

Public Sub CloseACARSConnection(Optional blnSendQuit As Boolean = False)
    blnUserQuit = False
    On Error Resume Next
    wsckMain.Close
    On Error GoTo 0
    DoEvents
    
    lstPilots.Clear
    config.ACARSConnected = False
    tmrRequest.Enabled = False
    blnACARSConnectionResolved = True
    sbMain.Panels(1).Text = "Status: Offline"
    mnuConnect.Caption = "Connect"
    cmdConnectDisconnect.Caption = "Connect"
    cmdConnectDisconnect.Enabled = True
    
    frmMain.txtPilotID.Enabled = True
    frmMain.txtPassword.Enabled = True
End Sub

Public Sub ProcessServerData(strData As String)
    Static strInputBuffer As String
    Dim intPos As Long

    'Ignore initial HELLO string.
    If (Not config.SeenHELO) And (InStr(1, strData, "DVA ACARS", vbTextCompare) > 0) And (InStr(1, strData, "HELLO", vbTextCompare) > 0) Then
        config.SeenHELO = True
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
    If InStr(strInput, ".") <> 1 Then
        SendChat strInput
        ReqStack.Send
    Else
        ProcessUserCmd strInput
    End If
End Sub

Private Sub ProcessUserCmd(strInput As String)
    Dim cmdName As String
    Dim aryParts As Variant
    
    strInput = Mid(strInput, 2, Len(strInput) - 1)
    aryParts = Split(strInput, " ", 2, vbTextCompare)
    cmdName = aryParts(0)

    'Process command accordingly.
    Select Case cmdName
        Case "msg"
            'Make sure we're connected
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
                Exit Sub
            End If
        
            If UBound(aryParts) < 1 Then
                ShowMessage "No message specified!", ACARSERRORCOLOR
                Exit Sub
            End If
                
            aryParts = Split(aryParts(1), " ", 2, vbTextCompare)

            'Make sure a message was specified.
            If UBound(aryParts) < 1 Then
                ShowMessage "No message specified!", ACARSERRORCOLOR
                Exit Sub
            End If

            'Process the chat message.
            SendChat CStr(aryParts(1)), CStr(aryParts(0))
            ReqStack.Send
            
        Case "pvtvoice"
            'Make sure we're online
            If Not config.SB3Support Then Exit Sub
            If Not config.SB3Connected Then
                ShowMessage "Squawkbox 3 not connected to FSD Server", ACARSERRORCOLOR
                Exit Sub
            ElseIf (config.PrivateVoiceURL = "") Then
                ShowMessage "No Private Voice URL", ACARSERRORCOLOR
                Exit Sub
            End If
            
            SB3PrivateVoice config.PrivateVoiceURL
            ShowMessage "Private Voice Channel set to " + config.PrivateVoiceURL, ACARSTEXTCOLOR

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
            
            'Make sure we're connected
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
            
            'Make sure we're connected
            If Not config.ACARSConnected Then
                ShowMessage "Not Connected to ACARS Server", ACARSERRORCOLOR
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
            
        Case "help"
            ShowMessage "Delta Virtual Airlines ACARS Help", ACARSTEXTCOLOR
            ShowMessage ".msg [userid] <msg> - Sends a message to another user", ACARSTEXTCOLOR
            ShowMessage ".pvtvoice - Tunes to DVA Private Voice Channel", ACARSTEXTCOLOR
            ShowMessage ".nav1 <vor> <heading> - Tunes NAV1 radio to a VOR", ACARSTEXTCOLOR
            ShowMessage ".nav2 <vor> <heading> - Tunes NAV2 radio to a VOR", ACARSTEXTCOLOR
            ShowMessage ".com1 <frequency> - Tunes COM1 radio to a frequency", ACARSTEXTCOLOR
            ShowMessage ".runway <airport> <runway> - Loads runway data/tunes ILS (if present)", ACARSTEXTCOLOR
            ShowMessage ".charts <airport> - Loads approach charts", ACARSTEXTCOLOR
            ShowMessage ".help - Display this help screen", ACARSTEXTCOLOR
            
        Case Else
            ShowMessage "Unknown command: " & aryParts(0), ACARSERRORCOLOR

    End Select
End Sub


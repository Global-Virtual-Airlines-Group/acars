Attribute VB_Name = "modMain"
Option Explicit

'API functions.
Declare Function GetForegroundWindow Lib "user32.dll" () As Long

'Constants for playing sounds.
Private Const SND_ALIAS = &H10000:
Private Const SND_FILENAME = &H20000:
Private Const SND_ASYNC = &H1:
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'ACARS constants.
Public Const ACARSVERSION = 1

'FSUIPC constants.
Public Const FSUIPCKEY = ""
Public Const FSUIPCREGOFFSET = &H8001&

'Message color constants.
Public Const ACARSTEXTCOLOR = &H109900
Public Const DEBUGTEXTCOLOR = &HAAAAAA
Public Const SELFCHATCOLOR = &H7B
Public Const PUBLICCHATCOLOR = vbBlack
Public Const PRIVATECHATCOLOR = vbRed
Public Const ACARSERRORCOLOR = vbRed
Public Const SYSMSGCOLOR = vbBlue

'Miscellaneous constants.
Public Const MAXTIMECOMPRESSION = 4
Public Const MINTIMECOMPRESSION = 1

'State variables
Public config As Configuration
Public info As New FlightData
Public pos As PositionData
Public positions As New OfflinePositionData

'Flight status variables.
Public FSUIPCErrors(15) As String
Public strFSUIPCErrorDesc As String
Public intFSUIPCErrorNum As Integer
Public intRetryFSConnection As Integer

Sub Main()
    frmSplash.Show
    frmMain.Icon = frmSplash.Icon

    'Make sure we're not already running.
    If App.PrevInstance = True Then
        MsgBox App.ProductName + " is already running!", vbOKOnly Or vbExclamation, "Error"
        End
    End If
    
    'Load configuration
    Set config = New Configuration
    config.LoadAirports
    config.LoadEquipment
    
    'Check for existing flight
    Dim oldInfo As FlightData
    Set oldInfo = config.LoadFlightInfo
    If (oldInfo.FlightID <> 0) Then
        If (MsgBox("You apparently had a flight in progress. Do you want to resume it?", _
            vbYesNo + vbExclamation, "Resume Flight") = vbYes) Then
            info.FlightID = oldInfo.FlightID
            info.startTime = oldInfo.startTime
        Else
            config.SaveFlightInfo 0
        End If
    End If
    
    'Update settings
    SetComboChoices frmMain.cboEquipment, config.EquipmentTypes, info.EquipmentType
    SetComboChoices frmMain.cboAirportD, config.AirportNames
    SetComboChoices frmMain.cboAirportA, config.AirportNames
    SetComboChoices frmMain.cboAirportL, config.AirportNames
    config.UpdateSettingsMenu
    config.UpdateFlightInfo
    
    'Set startup values for globals.
    config.SeenHELO = False
    FSUIPCErrors(1) = "FSUIPC: Attempt to Open when already Open"
    FSUIPCErrors(2) = "Cannot link to FSUIPC or WideClient"
    FSUIPCErrors(3) = "FSUIPC: Failed to Register common message with Windows"
    FSUIPCErrors(4) = "FSUIPC: Failed to create Atom for mapping filename"
    FSUIPCErrors(5) = "FSUIPC: Failed to create a file mapping object"
    FSUIPCErrors(6) = "FSUIPC: Failed to open a view to the file map"
    FSUIPCErrors(7) = "Incorrect version of FSUIPC, or not FSUIPC"
    FSUIPCErrors(8) = "FSUIPC: Sim is not version requested"
    FSUIPCErrors(9) = "FSUIPC: Call cannot execute, link not Open"
    FSUIPCErrors(10) = "FSUIPC: Call cannot execute: no requests accumulated"
    FSUIPCErrors(11) = "FSUIPC: IPC timed out all retries"
    FSUIPCErrors(12) = "FSUIPC: IPC sendmessage failed all retries"
    FSUIPCErrors(13) = "FSUIPC: IPC request contains bad data"
    FSUIPCErrors(14) = "FSUIPC: Maybe running on WideClient, but FS not running on Server, or wrong FSUIPC"
    FSUIPCErrors(15) = "FSUIPC: Read or Write request cannot be added, memory for Process is full"
    
    'Set enabled state of various controls.
    SetButtonMenuStates
    frmMain.tmrPosUpdates.Enabled = False
    frmMain.tmrRequest.Enabled = False

    'Set startup text for status bar panels.
    frmMain.sbMain.Panels(1).Text = "Status: Not connected"
    frmMain.sbMain.Panels(2).Text = "Flight Phase: N/A"
    frmMain.sbMain.Panels(3).Text = "Last Position Report: N/A"
    frmMain.sbMain.Panels(4).Text = "Sim Time: N/A"
    
    'Pause for 2 seconds
    Dim totalWait As Integer
    While ((totalWait < 2250) And Not frmSplash.isClicked)
        totalWait = totalWait + 250
        DoEvents
        Sleep 250
    Wend
    
    'Unload the splash screen
    frmSplash.Hide
    Unload frmSplash

    'Display the main form.
    frmMain.Show
    ShowMessage App.Title & " version " & App.Major & "." & App.Minor & " (Build " & _
        App.Revision & ") started", ACARSTEXTCOLOR

    'Set focus in command line text box.
    frmMain.txtCmd.SetFocus
End Sub

Public Function ConfirmExit() As Boolean
    ConfirmExit = True

    'If we have enough data for a PIREP, but the user hasn't filed
    'one, confirm that the user wants to quit.
    If info.FlightData And Not info.PIREPFiled Then
        ConfirmExit = (MsgBox("You have not filed a PIREP for your flight. Are you sure you want to exit?", vbYesNo Or vbQuestion, "Confirm") = vbYes)
        Exit Function
    End If

    'If we're in the middle of a flight, make sure the user really wants to quit.
    If info.InFlight Then
        ConfirmExit = frmMain.StopFlight()
        Exit Function
    End If

    'Prompt user to confirm exit if connected to ACARS server.
    If config.ACARSConnected Then
        ConfirmExit = (MsgBox("Are you sure you want to exit?", vbYesNo Or vbQuestion, "Confirm") = vbYes)
    End If
End Function

Public Function FSUIPCConnect() As Boolean
    Dim ResultText As Variant
    Dim SimulationText As Variant
    Dim dwResult As Long
    ResultText = Array( _
        "Connected", _
        "Attempt to Open when already Open", _
        "Cannot link to FSUIPC or WideClient", _
        "Failed to Register common message with Windows", _
        "Failed to create Atom for mapping filename", _
        "Failed to create a file mapping object", _
        "Failed to open a view to the file map", _
        "Incorrect version of FSUIPC, or not FSUIPC", _
        "Sim is not version requested", _
        "Call cannot execute, link not Open", _
        "Call cannot execute: no requests accumulated", _
        "IPC timed out all retries", _
        "IPC sendmessage failed all retries", _
        "IPC request contains bad data", _
        "Maybe running on WideClient, but FS not running on Server, or wrong FSUIPC", _
        "Read or Write request cannot be added, memory for Process is full" _
    )

    FSUIPCConnect = False
    config.FSUIPCConnected = False

    'Initialize important vars for UIPC comms - call only once!
    FSUIPC_Initialization

    'Try to connect to FSUIPC (or WideFS)
    If FSUIPC_Open(SIM_ANY, dwResult) Then
        If Len(FSUIPCKEY) > 0 Then
            If FSUIPC_Write(FSUIPCREGOFFSET, 12, VarPtr(FSUIPCKEY), dwResult) Then
                If Not FSUIPC_Process(dwResult) Then
                    intFSUIPCErrorNum = 513
                    strFSUIPCErrorDesc = "FSUIPC registration failure: " & ResultText(dwResult)
                    Exit Function
                End If
            Else
                intFSUIPCErrorNum = 514
                strFSUIPCErrorDesc = "FSUIPC registration failure: " & ResultText(dwResult)
                Exit Function
            End If
        End If
        
        'Get FS Versions
        Dim FSNames As Variant
        FSNames = Array("?", "FS98", "FS2000", "CFS2", "CFS1", "?", "FS2002", "FS2004")
        
        'Get the Flight Simulator version
        Dim FSVer As Integer
        Call FSUIPC_Read(&H3308, 2, VarPtr(FSVer), dwResult)
        If Not FSUIPC_Process(dwResult) Then Exit Function
        
        'Log FS Version
        If config.ShowDebug Then ShowMessage "Connected to " & FSNames(FSVer), DEBUGTEXTCOLOR
            
        'Save FS Version
        info.FSVersion = FSVer
        config.FSUIPCConnected = True
        FSUIPCConnect = True
    Else
        intFSUIPCErrorNum = 515
        strFSUIPCErrorDesc = "FSUIPC connection failure: " & ResultText(dwResult) & vbCrLf & vbCrLf & "Ensure that Flight Simulator is running."
    End If
End Function

Public Sub FSError(lngErrorCode As Long)
    'Dim strError As String
    FSUIPC_Close
    config.FSUIPCConnected = False
    If info.InFlight Then
        If MsgBox("The following error occurred:" & vbCrLf & vbCrLf & "(" & lngErrorCode & ") " & FSUIPCErrors(lngErrorCode) & vbCrLf & vbCrLf & "Press the Retry button to attempt the FSUIPC connection after fixing the problem, or press the Cancel button to end your flight.", vbCritical Or vbRetryCancel, "Error") = vbRetry Then
            'Attempt to reconnect to FSUIPC.
            While Not config.FSUIPCConnected
                FSUIPCConnect
                If Not config.FSUIPCConnected Then If MsgBox("The following error occurred:" & vbCrLf & vbCrLf & "(" & intFSUIPCErrorNum & ") " & strFSUIPCErrorDesc & vbCrLf & vbCrLf & "Press the Retry button to attempt the FSUIPC connection after fixing the problem, or press the Cancel button to end your flight.", vbCritical Or vbRetryCancel, "Error") = vbCancel Then GoTo ABORT
            Wend
        End If
    Else
        GoTo ABORT
    End If

ExitSub:
    Exit Sub

ABORT:
    frmMain.StopFlight True
    Resume ExitSub
End Sub

Public Sub SetButtonMenuStates()
    If info.InFlight Then
        frmMain.cmdStartStopFlight.Caption = "End Flight"
        frmMain.mnuFlightStartFlight.Enabled = False
        frmMain.mnuFlightEndFlight.Enabled = True
        frmMain.mnuOpenFlightPlan.Enabled = False
    Else
        frmMain.cmdStartStopFlight.Caption = "Start Flight"
        frmMain.mnuFlightStartFlight.Enabled = True
        frmMain.mnuFlightEndFlight.Enabled = False
        frmMain.mnuOpenFlightPlan.Enabled = config.SB3Support
    End If
End Sub

Public Sub ShowMessage(Msg As String, color As Long)
    Dim tStamp As String
    If config.ShowTimestamps Then tStamp = "[" & Format(Now, "hh:nn:ss") & "] "

    With frmMain.rtfText
        .SelColor = color
        .SelStart = Len(.TextRTF)
        .SelLength = 0
        
        If Len(.Text) > 0 Then
            .SelText = vbCrLf & tStamp & Msg
        Else
            .SelText = tStamp & Msg
        End If
        
        'Move the scrollbar to the end.
        .SelStart = Len(.TextRTF)
        .SelLength = 0
    End With
End Sub

Public Sub SetComboChoices(combo As ComboBox, choices As Variant, Optional newValue As String)
    Dim oldValue As String
    Dim x As Integer

    'Save the old value
    oldValue = combo.List(combo.ListIndex)
    If (newValue <> "") Then oldValue = newValue
    
    'Clear and add the values
    combo.Clear
    combo.AddItem "< SELECT >"
    For x = 0 To UBound(choices)
        combo.AddItem choices(x)
        If (oldValue = choices(x)) Then combo.ListIndex = (x + 1)
    Next
End Sub

Public Sub PlaySoundFile(name As String)
    Dim fName As String
    
    'Check that the file exists
    fName = App.Path + "\" + name
    If (Dir(fName) <> "") Then
        PlaySound fName, 0, SND_ASYNC And SND_FILENAME
    Else
        PlaySound "SystemAsterisk", 0, SND_ASYNC And SND_ALIAS
    End If
End Sub

Public Sub PlaySoundAlias(ByVal alias As String)
    PlaySound alias, 0, SND_ASYNC And SND_ALIAS
End Sub

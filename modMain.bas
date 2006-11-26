Attribute VB_Name = "modMain"
Option Explicit

'API functions.
Declare Function GetForegroundWindow Lib "user32.dll" () As Long

'Constants for playing sounds.
Private Const SND_ALIAS = &H10000:
Private Const SND_FILENAME = &H20000:
Private Const SND_ASYNC = &H1:
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

'ACARS constants.
Public Const ACARSVERSION = 1

'Message color constants.
Public Const ACARSTEXTCOLOR = &H109900
Public Const DEBUGTEXTCOLOR = &HAAAAAA
Public Const DEBUGGAUCOLOR = &H9F9F9F
Public Const SELFCHATCOLOR = &H7B
Public Const PUBLICCHATCOLOR = vbBlack
Public Const PRIVATECHATCOLOR = vbRed
Public Const ACARSERRORCOLOR = vbRed
Public Const SYSMSGCOLOR = vbBlue

'Debug data color constants
Public Const XML_IN_COLOR = &HAAAAAA
Public Const XML_OUT_COLOR = &HCACACA

'Minimum FSUIPC version constants
Private Const MIN_FSUIPC_CODE = &H36000000
Private Const MIN_FSUIPC_VERSION = "3.600"

'Miscellaneous constants
Public Const MAXTIMECOMPRESSION = 4
Public Const MINTIMECOMPRESSION = 1
Public Const MAX_TEXT_MESSAGES = 5

'State variables
Public config As Configuration
Public info As New FlightData
Public acInfo As AircraftInfo
Public pos As PositionData
Public Positions As New OfflinePositionData

'Online user info
Public users As New UserList

Sub Main()
    'Make sure we're not already running.
    If App.PrevInstance = True Then
        MsgBox App.ProductName + " is already running!", vbOKOnly Or vbExclamation, "Error"
        End
    End If
    
    'Load the splash screen
    frmSplash.Show vbModeless
    frmSplash.Icon = LoadResPicture(101, vbResIcon)
    
    'Do startup
    ApplicationStartup
    
    'Unload the splash screen
    Unload frmSplash
End Sub

Public Sub ApplicationStartup()
    Dim oldFlight As SavedFlight
    
    'Load main form
    Load frmMain
    frmMain.Icon = frmSplash.Icon
    
    'Disable the tabs
    frmMain.SSTab1.TabEnabled(1) = False
    frmMain.SSTab1.TabVisible(2) = False
    frmMain.SSTab1.TabEnabled(2) = False
    
    'Load configuration
    frmSplash.SetProgressLabel "Loading application configuration"
    Set config = New Configuration
    config.LoadAirports
    config.LoadEquipment
    config.LoadAirlines
    
    'Set startup message
    ShowMessage App.Title & " version " & App.Major & "." & App.Minor & " (Build " & _
        App.Revision & ") started", ACARSTEXTCOLOR

    'Check for existing flight
    Dim oldID As String
    frmSplash.SetProgressLabel "Checking for persisted Flight"
    oldID = config.OldFlightID
    If (oldID <> "") Then
        Dim hasFLT As Boolean
        
        'Load the flight data
        frmSplash.SetProgressLabel "Restoring Saved Flight data"
        Set oldFlight = RestoreFlightData(oldID)
        
        'Make sure we have a saved flight or the saved flight is complete
        config.IsFS9 = True
        hasFLT = FileExists(config.FS9Files + "\" + "ACARS Flight " + oldID + ".FLT")
        If Not hasFLT And (Not (oldFlight Is Nothing)) Then
            hasFLT = oldFlight.FlightInfo.FlightData Or config.WideFSInstalled
        Else
            hasFLT = hasFLT And (Not (oldFlight Is Nothing))
        End If
        
        'Confirm that we want to restore the flight
        If hasFLT Then
            hasFLT = hasFLT And (MsgBox("You apparently had a Flight in progress. Do you want to resume it?" + _
            vbCrLf + vbCrLf + "(Make sure Flight Simulator is already started and an aircraft is loaded.)", _
            vbYesNo + vbExclamation, "Resume In-Progress Flight") = vbYes)
        End If
        
        'If we want to reload the flight
        If hasFLT Then
            Dim dwResult As Long, totalWait As Integer
            Dim fName As String
            
            'Make sure FS is running if the flight is not complete
            If Not oldFlight.FlightInfo.FlightData Then
                While Not IsFSRunning
                    If (MsgBox("Microsoft Flight Simulator has not been started.", _
                        vbExclamation + vbRetryCancel, "Flight Simulator NotStarted") = _
                        vbCancel) Then End
                Wend
            
                'Connect to FSUIPC - if FS9 not running, abort
                frmSplash.SetProgressLabel "Connecting to FSUIPC"
                FSUIPC_Connect
                If Not config.FSUIPCConnected Then End
            
                'Make sure FS9 is "ready to fly"
                While Not IsFSReady
                    If (MsgBox("Microsoft Flight Simulator is not ""Ready to Fly"". Please" + _
                    vbCrLf + "ensure that an aircraft is loaded and ready to fly!", _
                    vbExclamation + vbRetryCancel, "Flight Simulator Not Ready") = vbCancel) Then
                        FSUIPC_Close
                        End
                    End If
                Wend
            End If
            
            'Load the saved flight
            Set info = oldFlight.FlightInfo
            ShowMessage "Restored Flight ID " + SavedFlightID(info), ACARSTEXTCOLOR
            If oldFlight.HasData Then
                Dim x As Integer
                Dim pd As PositionData
                    
                For x = 0 To UBound(oldFlight.Positions)
                    Set pd = oldFlight.Positions(x)
                    Positions.AddPosition pd
                Next
                
                If config.ShowDebug Then ShowMessage "Restored " + CStr(x) + " position cache entries", _
                    DEBUGTEXTCOLOR
            End If
            
            'Load the flight into FS9
            If Not info.FlightData Then
                frmSplash.SetProgressLabel "Restoring Microsoft Flight Simulator"
                fName = "ACARS Flight " + oldID + Chr(0)
                Call FSUIPC_WriteS(&H3F04, Len(fName) + 1, fName, dwResult)
                Call FSUIPC_Write(&H3F00, 2, VarPtr(FLIGHT_LOAD), dwResult)
                If Not FSUIPC_Process(dwResult) Then
                    FSUIPC_Close
                    MsgBox "Error Restoring Flight", vbError + vbOKOnly, "FSUIPC Error"
                    End
                ElseIf config.ShowDebug Then
                    ShowMessage "Restored Flight via FSUIPC", DEBUGTEXTCOLOR
                End If
            
                'Wait until we are ready to fly
                DoEvents
                Sleep 500
                Do
                    Sleep 250
                    totalWait = totalWait + 250
                Loop Until ((totalWait > 45000) Or IsFSReady)
            
                'If we're not ready, then kill things
                If Not IsFSReady Then
                    FSUIPC_Close
                    MsgBox "Error Restoring Flight", vbError + vbOKOnly, "FSUIPC Error"
                    End
                End If
            
                'Load aircraft info
                Set acInfo = GetAircraftInfo()
            Else
                If info.TestFlight Then
                    MsgBox "ACARS has loaded your old Flight Data, and you can export it to Microsoft " & _
                        "Excel and Google Earth.", vbOKOnly + vbInformation, "Flight Restored"
                Else
                    MsgBox "ACARS has loaded your old Flight Data, and has enough information to file" & _
                      vbCrLf & "a Flight Report. Please connect to the ACARS server to file your PIREP.", _
                      vbOKOnly + vbInformation, "Flight Restored"
                End If
            End If
            
            'Reset button states
            With frmMain
                .LockFlightInfo False
                .chkCheckRide.Enabled = Not info.TestFlight
                .chkTrainFlight.Enabled = Not info.CheckRide
                .tmrPosUpdates.Enabled = info.InFlight
                .tmrFlightTime.Enabled = info.InFlight
                .tmrStartCheck.Enabled = False
                .cmdPIREP.visible = True
                If info.TestFlight Then
                    .cmdPIREP.Enabled = info.FlightData
                    .cmdPIREP.Caption = "Export Data"
                    .chkTrainFlight.value = 1
                    .chkTrainFlight.Enabled = False
                Else
                    .cmdPIREP.Enabled = info.FlightData And config.ACARSConnected
                    If info.CheckRide Then .chkCheckRide.value = 1
                End If
            End With
        Else
            'Prompt if the user wants to delete the flight
            If (MsgBox("Do you want to delete this partial Flight record?", vbYesNo + vbQuestion, _
                "Delete Partial Flight") = vbYes) Then
                config.SaveFlightCode ""
                DeleteSavedFlight oldID
            End If
        End If
    End If

    'Update settings
    frmSplash.SetProgressLabel "Loading airport/equipment options"
    SetComboChoices frmMain.cboEquipment, config.EquipmentTypes, info.EquipmentType, "-"
    SetComboChoices frmMain.cboAirline, config.AirlineNames, info.Airline.Name, "-"
    SetAirport frmMain.cboAirportD, config.AirportNames, info.airportD
    SetAirport frmMain.cboAirportA, config.AirportNames, info.AirportA
    SetAirport frmMain.cboAirportL, config.AirportNames, info.AirportL
    config.UpdateSettingsMenu
    config.UpdateFlightInfo
    
    'If FS is running, connect to FSUIPC
    If IsFSRunning And Not config.FSUIPCConnected Then
        frmSplash.SetProgressLabel "Connecting to FSUIPC"
        FSUIPC_Connect
        If config.FSUIPCConnected Then Set acInfo = GetAircraftInfo()
    End If
    
    'Update gauge status
    GAUGE_SetPhase info.FlightPhase, config.ACARSConnected
    
    'Set enabled state of buttons
    SetButtonMenuStates

    'Set startup text for status bar panels
    With frmMain.sbMain
        .Panels(1).Text = "Status: Not connected"
        If (oldFlight Is Nothing) Then
            .Panels(2).Text = "Phase: N/A"
            .Panels(3).Text = "Last Position Report: N/A"
            .Panels(4).Text = "Flight Time: N/A"
        Else
            .Panels(2).Text = "Phase: " & info.PhaseName
            .Panels(3).Text = "Last Position Report: " & Format(Now, "hh:mm:ss")
            .Panels(4).Text = "Flight Time: " + Format(info.UpdateFlightTime(1, 0), "hh:mm:ss")
        End If
    End With
    
    'If we do not have a flight but FS is running, go into Cold & Dark mode
    If (oldFlight Is Nothing) And config.FSUIPCConnected And config.ColdDark Then
        frmSplash.SetProgressLabel "Setting Cold and Dark Cockpit"
        ColdDarkCockpit True
    End If
    
    'Pause for 2 seconds
    frmSplash.ClearProgressLabel
    If (oldFlight Is Nothing) Then
        While ((totalWait < 2250) And Not frmSplash.isClicked)
            totalWait = totalWait + 150
            DoEvents
            Sleep 150
        Wend
    End If
    
    'Display the main form.
    With frmMain
        .Show
        .txtCmd.SetFocus
        .MousePointer = vbDefault
        .SSTab1.TabVisible(3) = config.ShowDebug
        .chkStealth.Enabled = config.HasRole("HR")
        .chkStealth.visible = config.HasRole("HR")
        If config.FSUIPCConnected Then .tmrStartCheck.Enabled = True
    End With
End Sub

Public Function ConfirmExit() As Boolean
    ConfirmExit = True
    If info.FlightData And Not info.PIREPFiled Then
        If info.TestFlight Then
            ConfirmExit = (MsgBox("You have not saved your flight data from this Training Flight. " & _
                "Are you sure you want to exit?", vbYesNo Or vbQuestion, "Confirm Exit") = vbYes)
        Else
            ConfirmExit = (MsgBox("You have not filed a Flight Report for your flight. " & _
                "Are you sure you want to exit?", vbYesNo Or vbQuestion, "Confirm") = vbYes)
        End If
        
        If Not ConfirmExit Then Exit Function
    End If

    'If we're in the middle of a flight, make sure the user really wants to quit.
    If info.InFlight Then
        ConfirmExit = frmMain.StopFlight()
        Exit Function
    End If

    'Prompt user to confirm exit if connected to ACARS server.
    If config.ACARSConnected Then ConfirmExit = (MsgBox("Are you sure you want to exit?", _
        vbYesNo Or vbQuestion, "Confirm") = vbYes)
End Function

Public Function FSUIPC_Connect(Optional showError As Boolean = False) As Integer
    Dim WideFS_Version As Long
    Dim dwResult As Long

    config.FSUIPCConnected = False
    config.WideFSConnected = False
    
    'Make sure FS is running
    If Not IsFSRunning Then
        If showError Then MsgBox "Microsoft Flight Simulator has not been started.", _
            vbCritical + vbOKOnly, "Flight Simulator Not Started"
        FSUIPC_Connect = 17
        Exit Function
    End If

    'Initialize important vars for UIPC comms
    FSUIPC_Initialization

    'Try to connect to FSUIPC (or WideFS)
    If Not FSUIPC_Open(SIM_ANY, dwResult) Then
        If showError Then MsgBox FSUIPC_Error(dwResult), vbOKOnly + vbExclamation, "FSUIPC Error"
        FSUIPC_Connect = dwResult
        Exit Function
    End If
        
    'Load FSUIPC registration status and FS Version
    Dim Flags As Integer
    Dim FSVer As Integer
    Call FSUIPC_Read(&H3308, 2, VarPtr(FSVer), dwResult)
    Call FSUIPC_Read(&H330C, 2, VarPtr(Flags), dwResult)
    Call FSUIPC_Read(&H3322, 2, VarPtr(WideFS_Version), dwResult)
    If Not FSUIPC_Process(dwResult) Then
        MsgBox "Error reading FSUIPC Registration!", vbOKOnly Or vbCritical, "FSUIPC Error"
        FSUIPC_Close
        FSUIPC_Connect = 16
        Exit Function
    End If

    'Display FSUIPC registration status
    If ((Flags And 2) <> 0) Then
        ShowMessage "Registered FSUIPC detected", ACARSTEXTCOLOR
    ElseIf ((Flags And 1) <> 0) Then
        ShowMessage "Supplied freeware key to FSUIPC", ACARSTEXTCOLOR
    Else
        Dim EXEName As String
    
        'Display warning
        EXEName = LoadResString(102)
        ShowMessage "Unregistered FSUIPC access detected, expected " & EXEName & ".EXE, got " & _
            App.EXEName & ".EXE", ACARSERRORCOLOR
        FSUIPC_Close
        FSUIPC_Connect = 16
        Exit Function
    End If
    
    'Display WideFS
    If (WideFS_Version > 0) Then
        config.WideFSConnected = True
        If config.ShowDebug Then ShowMessage "WideClient detected", DEBUGTEXTCOLOR
    End If
    
    'Check that we're using a supported FSUIPC version
    If (FSUIPC_Version < MIN_FSUIPC_CODE) Then
        MsgBox App.ProductName & " requires FSUIPC v" & MIN_FSUIPC_VERSION & " or newer.", _
            vbOKOnly Or vbCritical, "Unsupported FSUIPC Version"
        FSUIPC_Close
        FSUIPC_Connect = 16
        Exit Function
    End If
    
    'Write ACARS status
    GAUGE_SetStatus ACARS_ON
        
    'Get FS Versions
    Dim FSNames As Variant
    FSNames = Array("?", "FS98", "FS2000", "CFS2", "CFS1", "?", "FS2002", "FS2004", "FSX")
        
    'Log FS Version
    If (FSVer > UBound(FSNames)) Then FSVer = UBound(FSNames)
    If config.ShowDebug Then ShowMessage "Connected to " & FSNames(FSVer), DEBUGTEXTCOLOR
            
    'Save FS Version
    info.FSVersion = FSVer
    config.IsFS9 = (FSVer >= 7)
    config.FSUIPCConnected = True
End Function

Public Sub FSError(ByVal errCode As Integer)
    Dim doCancel As Boolean

    'Close the FSUIPC connection
    FSUIPC_Close
    
    'If we're not flying then don't sweat it
    If Not info.InFlight Then Exit Sub
    
    'Check if FS is still running
    doCancel = Not IsFSRunning()
    
    'Prompt for reconnect
    While Not doCancel
        doCancel = (MsgBox("The following error occurred:" & vbCrLf & vbCrLf & _
        FSUIPC_Error(errCode) & vbCrLf & vbCrLf & _
        "Press the Retry button to reconnect to Microsoft" & vbCrLf & _
        "Flight Simulator, or Cancel to end your flight.", vbCritical Or vbRetryCancel, _
        "FSUIPC Error") = vbCancel)
        If Not doCancel Then
            errCode = FSUIPC_Connect(False)
            doCancel = doCancel Or config.FSUIPCConnected
        End If
    Wend
    
    'If we've not reconnected, cancel the flight without prejudice
    If Not config.FSUIPCConnected Then frmMain.StopFlight True
End Sub

Public Sub SetButtonMenuStates()
    frmMain.mnuOpenFlightPlan.Enabled = Not info.InFlight
    If info.InFlight Then
        frmMain.cmdStartStopFlight.Caption = "End Flight"
    ElseIf (info.FlightPhase = ERROR) Then
        frmMain.cmdStartStopFlight.Caption = "Recover Flight"
    Else
        frmMain.cmdStartStopFlight.Caption = "Start Flight"
    End If
End Sub

Public Sub ShowMessage(Msg As String, color As Long)
    Dim tStamp As String
    Dim oldSelPos As Long
    
    'Show timestamp if selected
    If config.ShowTimestamps Then tStamp = "[" & Format(Now, "hh:nn:ss") & "] "

    With frmMain.rtfText
        oldSelPos = .SelStart
        .SelStart = Len(.TextRTF)
        .SelLength = 0
        .SelColor = color
        If (Len(.Text) > 0) Then
            .SelText = vbCrLf & tStamp & Msg
        Else
            .SelText = tStamp & Msg
        End If
        
        'Move the scrollbar to the end.
        .SelLength = 0
        If config.FreezeWindow Then
            .SelStart = oldSelPos
        Else
            .SelStart = Len(.TextRTF)
        End If
    End With
End Sub

Public Sub ShowFSMessage(ByVal Msg As String, Optional scroll As Boolean = False, Optional wait As Integer = 8)
    Dim dwResult As Long
    Dim msgOptions As Integer

    'Ensure FSUIPC is connected
    If Not config.FSUIPCConnected Then Exit Sub
    
    'Truncate the message if necessary
    If (Len(Msg) > 127) Then Msg = Left(Msg, 127)
    Msg = Msg + Chr(0)
    
    'Set scrolling options
    If (wait < 1) Then wait = 2
    If scroll Then
        msgOptions = (wait * -1)
    Else
        msgOptions = wait
    End If
    
    'Write the message
    Call FSUIPC_WriteS(&H3380, Len(Msg), Msg, dwResult)
    Call FSUIPC_Write(&H32FA, 2, VarPtr(msgOptions), dwResult)
    If Not FSUIPC_Process(dwResult) Then ShowMessage "Error writing AdvMessage", ACARSERRORCOLOR
End Sub

Public Sub ShowDebug(XML As String, Optional color As Long = XML_IN_COLOR)
    With frmMain.rtfDebug
        .SelColor = color
        .SelStart = Len(.TextRTF)
        .SelLength = 0
        
        If Len(.Text) > 0 Then
            .SelText = vbCrLf & XML
        Else
            .SelText = XML
        End If
        
        'Move the scrollbar to the end.
        .SelStart = Len(.TextRTF)
        .SelLength = 0
    End With
End Sub

Public Sub SetAirport(combo As ComboBox, choices As Variant, ap As Airport)
    If (ap Is Nothing) Then
        SetComboChoices combo, choices, "-", "-"
    Else
        SetComboChoices combo, choices, ap.Name + " (" + ap.ICAO + ")", "-"
    End If
End Sub

Public Sub SetComboChoices(combo As ComboBox, choices As Variant, Optional newValue As String, Optional firstEntry As String)
    Dim oldValue As String
    Dim x As Integer
    Dim startOfs As Integer

    'Save the old value
    oldValue = combo.List(combo.ListIndex)
    If (newValue <> "") Then oldValue = newValue
    
    'Clear and add the first value
    combo.Clear
    If (firstEntry <> "") Then
        combo.AddItem firstEntry
        startOfs = 1
    End If
    
    'Add the remaining values
    For x = 0 To UBound(choices)
        combo.AddItem choices(x)
        If (oldValue = choices(x)) Then combo.ListIndex = (x + startOfs)
    Next
    
    'If we're not set, then set listindex to 0
    If (combo.ListIndex = -1) Then combo.ListIndex = 0
End Sub

Public Sub PlaySoundFile(Name As String)
    Dim fName As String
    
    'Check that the file exists
    fName = App.path + "\" + Name
    If (Dir(fName) <> "") Then
        PlaySound fName, 0, SND_ASYNC And SND_FILENAME
    Else
        PlaySound "SystemAsterisk", 0, SND_ASYNC And SND_ALIAS
    End If
End Sub

Public Sub PlaySoundAlias(ByVal alias As String)
    PlaySound alias, 0, SND_ASYNC And SND_ALIAS
End Sub

Public Sub LimitLength(txt As TextBox, ByVal maxLen As Integer, Optional doUpperCase As Boolean = False)
    If (Len(txt.Text) > maxLen) Then txt.Text = Left(txt.Text, maxLen)
    If (doUpperCase And (UCase(txt.Text) <> txt.Text)) Then
        txt.Text = UCase(txt.Text)
        txt.SelStart = Len(txt.Text)
    End If
End Sub

Public Sub LimitNumber(txt As TextBox, ByVal maxValue As Long)
    Dim x As Long
    Dim tst As String
    Dim c As String
    
    For x = 1 To Len(txt.Text)
        c = Mid(txt.Text, x, 1)
        If IsNumeric(c) Or (c = "-") Then tst = tst + c
    Next
    
    x = 0
    If ((tst <> "") And (tst <> "-")) Then x = CLng(tst)
    If (x > maxValue) Then x = maxValue
    txt.Text = CStr(x)
End Sub

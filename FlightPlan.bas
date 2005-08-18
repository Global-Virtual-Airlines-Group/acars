Attribute VB_Name = "SB3Support"
Option Explicit

Private Const SB3Filter = "Squawkbox 3 Flight Plans (*.sfp)|*.sfp"
Private Const FS9Filter = "FS2004 Flight Plans (*.pln)|*.pln"

Public Sub FPlan_Open()
    Dim planFileName As String
    
    If config.SB3Support Then
        frmMain.CommonDialog1.Filter = SB3Filter + "|" + FS9Filter
    Else
        frmMain.CommonDialog1.Filter = FS9Filter
    End If

    frmMain.CommonDialog1.CancelError = True
    frmMain.CommonDialog1.DialogTitle = "Open Squawkbox 3 Flight Plan"
    frmMain.CommonDialog1.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    
    'Display the dialog box
    On Error Resume Next
    frmMain.CommonDialog1.ShowOpen
    If err Then Exit Sub
    On Error GoTo 0
    
    'Get the file name
    planFileName = frmMain.CommonDialog1.FileName
    
    'Update the fields
    Select Case LCase(Right(planFileName, 3))
        Case "sfp"
            info.AirportD = ReadINI("SBFlightPlan", "Departure", info.AirportD, planFileName)
            info.AirportA = ReadINI("SBFlightPlan", "Arrival", info.AirportA, planFileName)
            info.CruiseAltitude = ReadINI("SBFlightPlan", "Altitude", info.CruiseAltitude, planFileName)
            info.Route = ReadINI("SBFlightPlan", "Route", info.Route, planFileName)
            info.Remarks = ReadINI("SBFlightPlan", "Remarks", info.Remarks, planFileName)
            
        Case "pln"
            Dim depInfo As String
            Dim destInfo As String
            Dim tmpRoute As String
            Dim x As Integer
            
            info.CruiseAltitude = ReadINI("flightplan", "cruising_altitude", info.CruiseAltitude, planFileName)
            depInfo = ReadINI("flightplan", "departure_id", "KATL", planFileName)
            destInfo = ReadINI("flightplan", "destination_id", "KATL", planFileName)
            info.AirportD = UCase(Trim(Split(depInfo, ",")(0)))
            info.AirportA = UCase(Trim(Split(destInfo, ",")(0)))
            
            While (depInfo <> "X")
                depInfo = ReadINI("flightplan", "waypoint." + CStr(x), "X", planFileName)
                If (depInfo <> "X") Then tmpRoute = tmpRoute + " " + UCase(Trim(Split(depInfo, ",")(1)))
                x = x + 1
            Wend
            
            info.Route = Trim(tmpRoute)
    End Select
    
    config.UpdateFlightInfo
End Sub

Public Sub SB3Plan_Save()
    Dim planFileName As String

    frmMain.CommonDialog1.CancelError = True
    frmMain.CommonDialog1.DialogTitle = "Save Squawkbox 3 Flight Plan"
    frmMain.CommonDialog1.Filter = "Squawkbox 3 Flight Plans (*.sfp)|*.sfp"
    frmMain.CommonDialog1.flags = cdlOFNHideReadOnly

    'Display the dialog box
    On Error Resume Next
    frmMain.CommonDialog1.ShowSave
    If err Then Exit Sub
    On Error GoTo 0
    
    'Get the file name
    planFileName = frmMain.CommonDialog1.FileName

    'Write the INI file
    WriteINI "SBFlightPlan", "Departure", info.AirportD, planFileName
    WriteINI "SBFlightPlan", "Arrival", info.AirportA, planFileName
    WriteINI "SBFlightPlan", "Altitude", info.CruiseAltitude, planFileName
    WriteINI "SBFlightPlan", "Route", info.Route, planFileName
    WriteINI "SBFlightPlan", "Remarks", info.Remarks, planFileName
End Sub

Public Function SB3InstallCheck() As Boolean
    Dim sb3Path As String
    
    sb3Path = RegReadString(HKEY_LOCAL_MACHINE, "SOFTWARE\Level 27 Technologies\Squawkbox\3", "SBDir", "")
    SB3InstallCheck = (sb3Path <> "")
End Function

Public Function SB3Running() As Boolean
    Dim lngResult As Long
    Dim appCode As Integer
    
    Call FSUIPC_Read(&H7B80, 1, VarPtr(appCode), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        FSError lngResult
        Exit Function
    End If
    
    SB3Running = (appCode > 0)
End Function

Public Function SB3Connected() As Boolean
    Dim lngResult As Long
    Dim appCode As Integer

    Call FSUIPC_Read(&H7B81, 1, VarPtr(appCode), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        FSError lngResult
        Exit Function
    End If
    
    SB3Connected = (appCode = 1)
End Function

Public Sub SB3Transponder(modeC As Boolean)
    Dim lngResult As Long
    Dim appCode As Integer

    If modeC Then
        appCode = 0
        ShowMessage "SB3 Transponder set to Mode C", ACARSTEXTCOLOR
    Else
        appCode = 1
        ShowMessage "SB3 Transponder set to Standby", ACARSTEXTCOLOR
    End If
    
    Call FSUIPC_Write(&H7B91, 1, VarPtr(appCode), lngResult)
    If Not FSUIPC_Process(lngResult) Then FSError lngResult
End Sub

Public Sub SB3PrivateVoice(ByVal url As String)
    Dim lngResult As Long
    Const appCode = 1

    url = url & Chr(0)
    Call FSUIPC_WriteS(&H7BA1, Len(url), url, lngResult)
    Call FSUIPC_Write(&H7BA0, 1, VarPtr(appCode), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        FSError lngResult
        Exit Sub
    End If
    
    ShowMessage "SB3 Private Voice channel set to " + url, ACARSTEXTCOLOR
End Sub

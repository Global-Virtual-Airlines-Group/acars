Attribute VB_Name = "FlightPlans"
Option Explicit

Private Const SB3Filter = "Squawkbox 3 Flight Plans (*.sfp)|*.sfp"
Private Const FS9Filter = "FS2004 Flight Plans (*.pln)|*.pln"
Private Const ACARSFilter = "ACARS Flight Data (ACARS Flight *.sha)|ACARS Flight *.sha"

Public Sub FPlan_Open()
    Dim oldPath As String
    Dim planFileName As String
    Dim lats As Variant
    Dim lngs As Variant
    Dim eqType As String
    
    'Set critical error handler
    On Error GoTo FatalError
    
    'Get equipment type
    If (Len(info.EquipmentType) < 4) Then
        MsgBox "Please select an Aircraft Type before loading a Flight Plan.", vbExclamation, _
            "No Aircraft Type"
        frmMain.cboEquipment.SetFocus
        Exit Sub
    End If
    
    'Set dialog box options
    With frmMain.CommonDialog1
        If config.SB3Support Then
            .Filter = FS9Filter + "|" + SB3Filter
        Else
            .Filter = FS9Filter
        End If
    
        oldPath = .InitDir
        If (config.SavedFlightsPath <> "") Then .InitDir = config.SavedFlightsPath
        .CancelError = True
        .DialogTitle = "Open Flight Plan"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        
        'Display the dialog box
        On Error Resume Next
        .ShowOpen
        .InitDir = oldPath
        If err Then Exit Sub
        On Error GoTo 0
        
        planFileName = .FileName
    End With
    
    'Save the flights path
    config.SavedFlightsPath = Left$(planFileName, Len(planFileName) - _
        Len(frmMain.CommonDialog1.FileTitle) - 1)
    
    'Update the fields
    Select Case LCase(Right(planFileName, 3))
        Case "sfp"
            Set info.airportD = config.GetAirport(ReadINI("SBFlightPlan", "Departure", "", planFileName))
            Set info.AirportA = config.GetAirport(ReadINI("SBFlightPlan", "Arrival", "", planFileName))
            Set info.AirportL = config.GetAirport(ReadINI("SBFlightPlan", "Alternate", "", planFileName))
            info.CruiseAltitude = ReadINI("SBFlightPlan", "Altitude", info.CruiseAltitude, planFileName)
            info.Route = ReadINI("SBFlightPlan", "Route", info.Route, planFileName)
            info.Remarks = ReadINI("SBFlightPlan", "Remarks", info.Remarks, planFileName)
            
        Case "pln"
            Dim depInfo As String
            Dim depData As Variant
            Dim destInfo As String
            Dim tmpRoute As String
            Dim routeType As Integer
            Dim x As Integer
            
            'Initialize the data arrays
            ReDim lats(25)
            ReDim lngs(25)
            
            routeType = CInt(ReadINI("flightplan", "routetype", "3", planFileName))
            info.CruiseAltitude = ReadINI("flightplan", "cruising_altitude", info.CruiseAltitude, planFileName)
            depInfo = ReadINI("flightplan", "departure_id", "KATL", planFileName)
            destInfo = ReadINI("flightplan", "destination_id", "KATL", planFileName)
            Set info.airportD = config.GetAirport(UCase(Trim(Split(depInfo, ",")(0))))
            Set info.AirportA = config.GetAirport(UCase(Trim(Split(destInfo, ",")(0))))
            
            While (depInfo <> "X")
                depInfo = ReadINI("flightplan", "waypoint." + CStr(x), "X", planFileName)
                If (depInfo <> "X") Then
                    depData = Split(UCase(depInfo), ",")
                    
                    'Parse different data
                    If ((routeType = 3) Or (routeType = 0)) Then
                        tmpRoute = tmpRoute + " " + Trim(depData(3))
                        If (x < 25) Then
                            lats(x) = ConvertLatLon(depData(5))
                            lngs(x) = ConvertLatLon(depData(6))
                        End If
                    Else
                        tmpRoute = tmpRoute + " " + Trim(depData(0))
                        If (x < 25) Then
                            lats(x) = ConvertLatLon(depData(2))
                            lngs(x) = ConvertLatLon(depData(3))
                        End If
                    End If
                End If
                
                x = x + 1
            Wend
            
            'Save the route
            info.Route = UCase(Trim(tmpRoute))
            
            'If we're using a 707/720, offer to write the flight plan
            eqType = Left(info.EquipmentType, 4)
            If ((eqType = "B707") Or (eqType = "B720")) Then
                If (MsgBox("Do you want to save this as a Boeing 707 INS flight plan?", _
                    vbYesNo + vbQuestion, "707 INS Flight Plan") = vbYes) Then
                        If (x < 25) Then
                            ReDim Preserve lats(x - 1)
                            ReDim Preserve lngs(x - 1)
                        Else
                            ShowMessage "Flight Plan truncated. INS plans have a maximum of 25 waypoints.", ACARSERRORCOLOR
                        End If
                        
                        'Save the flight plan
                        SaveINSPlan lats, lngs
                End If
            End If
    End Select
    
    config.UpdateFlightInfo
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Error #" & CStr(err.Number) & " (" & err.Description & ") loading Flight Plan", vbCritical, "Error"
    Resume ExitSub
    
End Sub

Public Sub SB3Plan_Save()
    Dim planFileName As String
    
    'Check that airports are set
    If ((info.airportD Is Nothing) Or (info.AirportA Is Nothing)) Then Exit Sub
    
    'Get the path/dialog options
    With frmMain.CommonDialog1
        .FileName = info.airportD.ICAO + "-" + info.AirportA.ICAO + ".sfp"
        .CancelError = True
        .DialogTitle = "Save Squawkbox 3 Flight Plan"
        .Filter = "Squawkbox 3 Flight Plans (*.sfp)|*.sfp"
        .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
    End With
    
    'Display the dialog box
    On Error Resume Next
    frmMain.CommonDialog1.ShowSave
    If err Then Exit Sub
    On Error GoTo FatalError
    
    'Get the file name
    planFileName = frmMain.CommonDialog1.FileName

    'Write the INI file
    If Not (info.airportD Is Nothing) Then WriteINI "SBFlightPlan", "Departure", info.airportD.ICAO, planFileName
    If Not (info.AirportA Is Nothing) Then WriteINI "SBFlightPlan", "Arrival", info.AirportA.ICAO, planFileName
    If Not (info.AirportL Is Nothing) Then WriteINI "SBFlightPlan", "Alternate", info.AirportL.ICAO, planFileName
    WriteINI "SBFlightPlan", "Altitude", info.CruiseAltitude, planFileName
    WriteINI "SBFlightPlan", "Route", info.Route, planFileName
    WriteINI "SBFlightPlan", "Remarks", info.Remarks, planFileName
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Error saving SquawkBox 3 Flight Plan", vbCritical, err.Description
    Resume ExitSub
    
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

Public Sub SB3PrivateVoice(ByVal Url As String)
    Dim lngResult As Long
    Const appCode = 1

    Url = Url & Chr(0)
    Call FSUIPC_WriteS(&H7BA1, Len(Url), Url, lngResult)
    Call FSUIPC_Write(&H7BA0, 1, VarPtr(appCode), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        FSError lngResult
        Exit Sub
    End If
    
    ShowMessage "SB3 Private Voice channel set to " + Url, ACARSTEXTCOLOR
End Sub

Private Function ConvertLatLon(ByVal info As String) As Double
    Dim degParts As Variant
    Dim degrees As Double
    Dim hemi As String
    
    'Remove leading space and split string
    degParts = Split(Mid(info, 2), "* ")
    
    'Get the hemisphere
    hemi = UCase(Left(degParts(0), 1))
    
    'Calculate degrees and minutes
    degrees = Int(Val(Mid(degParts(0), 2)))
    degrees = degrees + (Val(Left(degParts(1), Len(degParts(1)) - 1)) / 60)
    If ((hemi = "S") Or (hemi = "W")) Then degrees = degrees * -1
    
    'Combine and return
    ConvertLatLon = degrees
End Function

Private Sub SaveINSPlan(lats As Variant, lngs As Variant)
    Dim x As Integer
    Dim fNum As Integer

    With frmMain.CommonDialog1
        .CancelError = True
        .DialogTitle = "Save B707 INS Flight Plan"
        .Filter = "B707 INS Flight Plans (707fplan?.dat)|707fplan?.dat"
        .Flags = cdlOFNHideReadOnly
        .FileName = "707fplan0.dat"
    End With

    'Display the dialog box
    On Error Resume Next
    frmMain.CommonDialog1.ShowSave
    If err Then Exit Sub
    On Error GoTo 0
    
    'Write to the file
    fNum = FreeFile()
    Open frmMain.CommonDialog1.FileName For Output Lock Write As #fNum
    For x = 0 To UBound(lats)
        Print #fNum, Format(lats(x), "#0.000000")
        If (lngs(x) < 0) Then
            Print #fNum, Format(lngs(x) + 360, "##0.000000")
        Else
            Print #fNum, Format(lngs(x), "##0.000000")
        End If
    Next
    
    Close #fNum
End Sub

Public Sub FData_Open()
    Dim oldPath As String
    Dim shaFileName As String
    Dim FlightID As String
    Dim oldFlight As SavedFlight
    Dim oldInfo As FlightData

    'Get the file to load
    With frmMain.CommonDialog1
        oldPath = .InitDir
        .InitDir = App.path
        .Filter = ACARSFilter
        .CancelError = True
        .DialogTitle = "Open Flight Data"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
        
        'Display the dialog box
        On Error Resume Next
        .ShowOpen
        .InitDir = oldPath
        If err Then Exit Sub
        On Error GoTo 0
        
        shaFileName = .FileName
        FlightID = Left(.FileTitle, InStr(.FileTitle, ".") - 1)
        FlightID = Right(FlightID, Len(FlightID) - 13)
    End With
    
    'Load the flight
    Set oldFlight = RestoreFlightData(FlightID)
    If (oldFlight Is Nothing) Then
        MsgBox "Cannot Restore Flight " & FlightID & "!", vbExclamation Or vbOKOnly, "Cannot Load Data"
        Exit Sub
    End If
    
    'Get the flight data
    If config.ShowDebug Then ShowMessage "Loaded Flight " + FlightID, DEBUGTEXTCOLOR
    Set oldInfo = oldFlight.FlightInfo
    If Not oldInfo.FlightData Then
        MsgBox "Flight " & FlightID & " is not complete, and cannot be loaded.", vbExclamation Or _
            vbOKOnly, "Flight Not Completed"
        Exit Sub
    ElseIf oldInfo.TestFlight Then
        MsgBox "Flight " & FlightID & " was a Test Flight, and cannot be loaded.", vbExclamation Or _
            vbOKOnly, "Test Flight"
        Exit Sub
    End If
    
    'Load the position cache
    Dim x As Integer
    Dim pd As PositionData
    For x = 0 To UBound(oldFlight.Positions)
        Set pd = oldFlight.Positions(x)
        Positions.AddPosition pd
    Next
    
    If config.ShowDebug Then ShowMessage "Restored " + CStr(x) + " position cache entries", DEBUGTEXTCOLOR
        
    'Reset button states
    Set info = oldInfo
    config.UpdateFlightInfo
    With frmMain
        .LockFlightInfo False
        .chkCheckRide.Enabled = Not info.TestFlight
        .chkTrainFlight.Enabled = Not info.CheckRide
        .tmrPosUpdates.Enabled = info.InFlight
        .tmrFlightTime.Enabled = info.InFlight
        .tmrStartCheck.Enabled = False
        .cmdPIREP.visible = True
        .cmdPIREP.Enabled = info.FlightData And config.ACARSConnected
        If info.CheckRide Then .chkCheckRide.value = 1
    End With
    
    'Display status message
    MsgBox "ACARS has loaded your old Flight Data, and has enough information" & vbCrLf & _
        "to file a Flight Report.", vbOKOnly Or vbInformation, "Flight Restored"
End Sub

Attribute VB_Name = "SB3Support"
Option Explicit

Private Const SB3Filter = "Squawkbox 3 Flight Plans (*.sfp)|*.sfp"
Private Const FS9Filter = "FS2004 Flight Plans (*.pln)|*.pln"

Public Sub FPlan_Open()
    Dim planFileName As String
    Dim lats As Variant
    Dim lngs As Variant
    
    'Set dialog box options
    With frmMain.CommonDialog1
        If config.SB3Support Then
            .Filter = FS9Filter + "|" + SB3Filter
        Else
            .Filter = FS9Filter
        End If
    
        If (config.SavedFlightsPath <> "") Then .InitDir = config.SavedFlightsPath
        .CancelError = True
        .DialogTitle = "Open Flight Plan"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    End With
    
    'Display the dialog box
    On Error Resume Next
    frmMain.CommonDialog1.ShowOpen
    If err Then Exit Sub
    On Error GoTo 0
    
    'Get the file name and save the flights path
    planFileName = frmMain.CommonDialog1.FileName
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
            Dim x As Integer
            
            'Initialize the data arrays
            ReDim lats(25)
            ReDim lngs(25)
            
            info.CruiseAltitude = ReadINI("flightplan", "cruising_altitude", info.CruiseAltitude, planFileName)
            depInfo = ReadINI("flightplan", "departure_id", "KATL", planFileName)
            destInfo = ReadINI("flightplan", "destination_id", "KATL", planFileName)
            Set info.airportD = config.GetAirport(UCase(Trim(Split(depInfo, ",")(0))))
            Set info.AirportA = config.GetAirport(UCase(Trim(Split(destInfo, ",")(0))))
            
            While (depInfo <> "X")
                depInfo = ReadINI("flightplan", "waypoint." + CStr(x), "X", planFileName)
                If (depInfo <> "X") Then
                    depData = Split(UCase(depInfo), ",")
                    tmpRoute = tmpRoute + " " + Trim(depData(3))
                    If (x < 25) Then
                        lats(x) = ConvertLatLon(depData(5))
                        lngs(x) = ConvertLatLon(depData(6))
                    End If
                End If
                
                x = x + 1
            Wend
            
            'Save the route
            info.Route = UCase(Trim(tmpRoute))
            
            'If we're using a 707/720, offer to write the flight plan
            Dim eqType As String
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
End Sub

Public Sub SB3Plan_Save()
    Dim planFileName As String
    
    'Get the path/dialog options
    With frmMain.CommonDialog1
        .FileName = frmMain.CommonDialog1.InitDir + info.airportD.ICAO + "-" + info.AirportA.ICAO + ".sfp"
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

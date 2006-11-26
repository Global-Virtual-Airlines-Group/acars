Attribute VB_Name = "RequestSender"
Option Explicit

Public ReqStack As New RequestStack

Private Function buildCMD(cmdType As String) As IXMLDOMElement
    Dim doc As New DOMDocument
    Dim cmdE As IXMLDOMElement
   
    'Create the CMD node
    Set cmdE = doc.createNode(NODE_ELEMENT, "CMD", "")
    cmdE.setAttribute "type", cmdType

    'Return the document
    Set buildCMD = cmdE
End Function

Public Sub RequestPilotList()
    Dim cmd As IXMLDOMElement
    
    'Send the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "pilots", False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent pilot list request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestPrivateVoiceURL()
    Dim cmd As IXMLDOMElement
    
    'Send the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "pvtvox", False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent private voice URL request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestDraftPIREPs()
    Dim cmd As IXMLDOMElement
    
    'Send the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "draftpirep", False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent draft Flight Report request", DEBUGTEXTCOLOR
End Sub

Public Function SendFlightInfo(fInfo As FlightData) As Long
    Dim cmd As IXMLDOMElement

    'Build the request and save the number
    Set cmd = buildCMD("flight_info")
    
    'Add flight plan info
    AddXMLField cmd, "flight_num", fInfo.flightCode, False
    AddXMLField cmd, "leg", CStr(fInfo.FlightLeg), False
    AddXMLField cmd, "equipment", fInfo.EquipmentType, False
    AddXMLField cmd, "cruise_alt", fInfo.CruiseAltitude, False
    AddXMLField cmd, "airportD", fInfo.airportD.IATA, False
    AddXMLField cmd, "airportA", fInfo.AirportA.IATA, False
    AddXMLField cmd, "route", fInfo.Route, True
    AddXMLField cmd, "remarks", fInfo.Remarks, True
    AddXMLField cmd, "startTime", FormatDateTime(fInfo.StartTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "fs_ver", CStr(fInfo.FSVersion), False
    If fInfo.Offline Then AddXMLField cmd, "offline", "true", False
    If (fInfo.FlightPhase = COMPLETE) Then AddXMLField cmd, "complete", "true", False
    If (fInfo.FlightID > 0) Then
        AddXMLField cmd, "flight_id", CStr(fInfo.FlightID), False
        ShowMessage "Resuming Flight " + CStr(fInfo.FlightID), ACARSTEXTCOLOR
    End If
    
    'Queue the request
    ReqStack.Queue cmd
    SendFlightInfo = ReqStack.RequestID
    
    If config.ShowDebug Then ShowMessage "Sent flight info " & Hex(SendFlightInfo), DEBUGTEXTCOLOR
End Function

Public Sub RequestPilotInfo(PilotID As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim Flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "pilot"

    'Set the pilot ID in the flags
    Set Flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild Flags
    AddXMLField Flags, "pilot_id", PilotID, False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent pilot info request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestATCInfo(Network As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim Flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "atc", False

    'Set the network name in the flags
    Set Flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild Flags
    AddXMLField Flags, "network", Network, False
    ReqStack.Queue cmd

    If config.ShowDebug Then ShowMessage "Sent " + Network + " ATC info request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestRunwayInfo(AirportCode As String, runwayCode As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim Flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "navaid"

    'Set the navaid info in the flags
    Set Flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild Flags
    AddXMLField Flags, "id", AirportCode, False
    AddXMLField Flags, "runway", runwayCode, False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent Runway Info request", DEBUGTEXTCOLOR
End Sub

Public Sub SendBusy(Optional ImBusy As Boolean = True)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim Flags As IXMLDOMElement
    
    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "busy", False

    'Set the navaid info in the flags
    Set Flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild Flags
    AddXMLField Flags, "isBusy", CStr(ImBusy), False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Toggled Busy status", DEBUGTEXTCOLOR
End Sub

Public Sub RequestNavaidInfo(navaidName As String, hdg As String, radioName As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim Flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "navaid", False

    'Set the navaid info in the flags
    Set Flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild Flags
    AddXMLField Flags, "id", navaidName, False
    AddXMLField Flags, "radio", radioName, False
    AddXMLField Flags, "hdg", hdg, False
    ReqStack.Queue cmd

    If config.ShowDebug Then ShowMessage "Sent Navigation Aid request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestCharts(AirportCode As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim Flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "charts", False

    'Set the navaid info in the flags
    Set Flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild Flags
    AddXMLField Flags, "id", UCase(AirportCode), False
    ReqStack.Queue cmd

    If config.ShowDebug Then ShowMessage "Sent Approach Chart request", DEBUGTEXTCOLOR
End Sub

Public Sub SendChat(msgText As String, Optional msgTo As String)
    Dim cmd As IXMLDOMElement
    Dim msgFrom As String
    Dim p As Pilot

    'Build the request and add the message text
    Set cmd = buildCMD("text")
    AddXMLField cmd, "text", msgText, True
    
    'Build the originator string
    If config.ShowPilotNames Then
        msgFrom = frmMain.txtPilotName.Text
    Else
        msgFrom = frmMain.txtPilotID.Text
    End If
    
    'Find the recipient
    If (Len(msgTo) > 0) Then
        Set p = users.GetPilot(msgTo)
        If (p Is Nothing) Then
            ShowMessage msgTo & " is not logged in!", ACARSERRORCOLOR
            Exit Sub
        ElseIf (p.ID = frmMain.txtPilotID.Text) Then
            Exit Sub
        End If
    
        AddXMLField cmd, "to", p.ID
        If config.ShowPilotNames Then
            msgFrom = msgFrom + "->" + p.Name
        Else
            msgFrom = msgFrom + "->" + p.ID
        End If
    End If

    'Send the message
    ReqStack.Queue cmd
    
    'Show outgoing message
    ShowMessage "<" & msgFrom & "> " & msgText, SELFCHATCOLOR
    If config.ShowDebug Then ShowMessage "Sent chat message", DEBUGTEXTCOLOR
End Sub

Public Function SendCredentials(userID As String, pwd As String) As Long
    Dim cmd As IXMLDOMElement
    Dim isStealth As Boolean
    
    'Log stealth mode
    isStealth = (frmMain.chkStealth.value = 1) And config.HasRole("HR")
    If isStealth And config.ShowDebug Then ShowMessage "Hidden/Stealth Connection", DEBUGTEXTCOLOR

    'Build the request
    Set cmd = buildCMD("auth")

    'Add user and password
    AddXMLField cmd, "user", userID, False
    AddXMLField cmd, "password", pwd, True
    AddXMLField cmd, "build", CStr(App.Revision), False
    AddXMLField cmd, "stealth", CStr(isStealth), False
    AddXMLField cmd, "version", "v" & CStr(App.Major) & "." & CStr(App.Minor), False
    ReqStack.Queue cmd
    SendCredentials = ReqStack.RequestID

    If config.ShowDebug Then ShowMessage "Logging In", DEBUGTEXTCOLOR
End Function

Public Sub SendPing()
    Dim cmd As IXMLDOMElement

    'Build the request and send it
    ReqStack.Queue buildCMD("ping")
    If config.ShowDebug Then ShowMessage "Sent ping", DEBUGTEXTCOLOR
End Sub

Public Function SendEndFlight() As Long
    Dim cmd As IXMLDOMElement

    'Build the request and send it
    ReqStack.Queue buildCMD("end_flight")
    If config.ShowDebug Then ShowMessage "Sent EndFlight " & Hex(ReqStack.RequestID), DEBUGTEXTCOLOR
    SendEndFlight = ReqStack.RequestID
End Function

Public Sub RequestEquipment()
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "eqList", False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent equipment list request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestAirlines()
    Dim cmd As IXMLDOMElement
    
    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "aList", False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent airline list request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestAirports()
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "apList", False
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent airport list request", DEBUGTEXTCOLOR
End Sub

Public Sub SendKick(ByVal PilotID As String, Optional banAddr As Boolean = False)
    Dim cmd As IXMLDOMElement
    
    'Build the request
    Set cmd = buildCMD("diag")
    AddXMLField cmd, "reqdata", UCase(PilotID)
    If banAddr Then
        AddXMLField cmd, "reqtype", "BlockIP", False
    Else
        AddXMLField cmd, "reqtype", "KickUser", False
    End If
    
    ReqStack.Queue cmd
    If config.ShowDebug Then ShowMessage "Sent kick request", DEBUGTEXTCOLOR
End Sub

Public Function SendPosition(ByVal cPos As PositionData, ByVal IsLogged As Boolean, _
    ByVal noFlood As Boolean) As Long
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("position")
    
    'Add position info nodes.
    AddXMLField cmd, "lat", FormatNumber(cPos.Latitude, "#0.00000"), False
    AddXMLField cmd, "lon", FormatNumber(cPos.Longitude, "##0.00000"), False
    AddXMLField cmd, "msl", cPos.AltitudeMSL, False
    AddXMLField cmd, "agl", cPos.AltitudeAGL, False
    AddXMLField cmd, "hdg", cPos.Heading, False
    AddXMLField cmd, "aSpeed", cPos.Airspeed, False
    AddXMLField cmd, "gSpeed", cPos.GroundSpeed, False
    AddXMLField cmd, "vSpeed", cPos.VerticalSpeed, False
    AddXMLField cmd, "aoa", FormatNumber(cPos.AngleOfAttack, "#0.000"), False
    AddXMLField cmd, "g", FormatNumber(cPos.GForce, "#0.000"), False
    AddXMLField cmd, "pitch", FormatNumber(cPos.Pitch, "#0.000"), False
    AddXMLField cmd, "bank", FormatNumber(cPos.Bank, "#0.000"), False
    AddXMLField cmd, "mach", FormatNumber(cPos.Mach, "0.000"), False
    AddXMLField cmd, "n1", FormatNumber(cPos.AverageN1, "##0.0"), False
    AddXMLField cmd, "n2", FormatNumber(cPos.AverageN2, "##0.0"), False
    AddXMLField cmd, "fuelFlow", CStr(cPos.FuelFlow), False
    AddXMLField cmd, "phase", cPos.phase, False
    AddXMLField cmd, "simrate", CStr(cPos.simRate / 256), False
    AddXMLField cmd, "flaps", CStr(cPos.Flaps), False
    AddXMLField cmd, "fuel", CStr(cPos.Fuel), False
    AddXMLField cmd, "weight", CStr(cPos.weight), False
    AddXMLField cmd, "wHdg", CStr(cPos.WindHeading), False
    AddXMLField cmd, "wSpeed", CStr(cPos.WindSpeed), False
    AddXMLField cmd, "date", FormatDateTime(cPos.DateTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "flags", CStr(cPos.Flags), False
    AddXMLField cmd, "frameRate", CStr(cPos.FrameRate), False
    If noFlood Then AddXMLField cmd, "noFlood", "true", False
    If Not IsLogged Then AddXMLField cmd, "isLogged", "false", False

    'Send the request
    ReqStack.Queue cmd
    SendPosition = ReqStack.RequestID
    
    'Log the request
    If config.ShowDebug Then
        If noFlood Then
            ShowMessage "Sent bulk position update " & Hex(SendPosition), DEBUGTEXTCOLOR
        Else
            ShowMessage "Sent position update " & Hex(SendPosition), DEBUGTEXTCOLOR
        End If
    End If
End Function

Public Function SendPIREP(info As FlightData) As Long
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("pirep")

    'Add pirep info fields
    AddXMLField cmd, "flightID", CStr(info.FlightID), False
    If info.CheckRide Then AddXMLField cmd, "checkRide", "true", False
    AddXMLField cmd, "flightcode", info.flightCode, False
    AddXMLField cmd, "leg", CStr(info.FlightLeg), False
    AddXMLField cmd, "eqType", info.EquipmentType, False
    AddXMLField cmd, "airportD", info.airportD.IATA, False
    AddXMLField cmd, "airportA", info.AirportA.IATA, False
    AddXMLField cmd, "remarks", info.Remarks, True
    AddXMLField cmd, "network", info.Network, False
    AddXMLField cmd, "startTime", FormatDateTime(info.StartTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "taxiOutTime", FormatDateTime(info.TaxiOutTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "taxiFuel", CStr(info.TaxiFuel), False
    AddXMLField cmd, "taxiWeight", CStr(info.TaxiWeight), False
    AddXMLField cmd, "takeoffTime", FormatDateTime(info.TakeoffTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "takeoffFuel", CStr(info.TakeoffFuel), False
    AddXMLField cmd, "takeoffWeight", CStr(info.TakeoffWeight), False
    AddXMLField cmd, "takeoffN1", FormatNumber(info.TakeoffN1, "##0.0")
    AddXMLField cmd, "takeoffSpeed", CStr(info.TakeoffSpeed), False
    AddXMLField cmd, "landingTime", FormatDateTime(info.LandingTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "landingFuel", CStr(info.LandingFuel), False
    AddXMLField cmd, "landingWeight", CStr(info.LandingWeight), False
    AddXMLField cmd, "landingN1", FormatNumber(info.LandingN1, "##0.0")
    AddXMLField cmd, "landingSpeed", CStr(info.LandingSpeed), False
    AddXMLField cmd, "landingVSpeed", CStr(info.LandingVSpeed), False
    AddXMLField cmd, "gateTime", FormatDateTime(info.GateTime.UTCTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "gateFuel", CStr(info.GateFuel), False
    AddXMLField cmd, "gateWeight", CStr(info.GateWeight), False
    AddXMLField cmd, "time0X", CStr(info.TimePaused), False
    AddXMLField cmd, "time1X", CStr(info.TimeAt1X), False
    AddXMLField cmd, "time2X", CStr(info.TimeAt2X), False
    AddXMLField cmd, "time4X", CStr(info.TimeAt4X), False

    'Send the request
    ReqStack.Queue cmd
    SendPIREP = ReqStack.RequestID
    If config.ShowDebug Then ShowMessage "Sent Flight Report " & Hex(SendPIREP), DEBUGTEXTCOLOR
End Function

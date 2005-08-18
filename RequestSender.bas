Attribute VB_Name = "RequestSender"
Option Explicit

Public ReqStack As New RequestStack

Private Function BeautifyXML(strXML As String) As String
    BeautifyXML = Replace(strXML, "><", ">" & vbCrLf & "<")
End Function

Private Function buildCMD(cmdType As String) As IXMLDOMElement
    Dim doc As New DOMDocument
    Dim cmdE As IXMLDOMElement
   
    'Create the CMD node
    Set cmdE = doc.createNode(NODE_ELEMENT, "CMD", "")
    cmdE.setAttribute "type", cmdType

    'Return the document
    Set buildCMD = cmdE
End Function

Private Sub AddXMLField(objNode As IXMLDOMElement, name As String, value As String)
    Dim objNewNode As IXMLDOMElement
    Dim objCDATA As IXMLDOMCDATASection

    Set objCDATA = objNode.ownerDocument.createCDATASection(value)
    Set objNewNode = objNode.ownerDocument.createNode(NODE_ELEMENT, name, "")
    objNewNode.appendChild objCDATA
    objNode.appendChild objNewNode
End Sub

Public Sub RequestPilotList()
    Dim cmd As IXMLDOMElement
    
    'Send the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "pilots"
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent pilot list request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestPrivateVoiceURL()
    Dim cmd As IXMLDOMElement
    
    'Send the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "pvtvox"
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent private voice URL request", DEBUGTEXTCOLOR
End Sub

Public Sub SendFlightInfo(fInfo As FlightData)
    Dim cmd As IXMLDOMElement

    'Build the request and save the number
    Set cmd = buildCMD("flight_info")
    fInfo.InfoReqID = ReqStack.RequestID
    
    'Add flight plan info
    AddXMLField cmd, "flight_num", fInfo.FlightNumber
    AddXMLField cmd, "equipment", fInfo.EquipmentType
    AddXMLField cmd, "cruise_alt", fInfo.CruiseAltitude
    AddXMLField cmd, "airportD", fInfo.AirportD
    AddXMLField cmd, "airportA", fInfo.AirportA
    AddXMLField cmd, "route", fInfo.Route
    AddXMLField cmd, "remarks", fInfo.Remarks
    AddXMLField cmd, "fs_ver", CStr(fInfo.FSVersion)
    If fInfo.Offline Then AddXMLField cmd, "offline", "true"
    If (fInfo.Phase = "Completed") Then AddXMLField cmd, "complete", "true"
    If (fInfo.flightID > 0) Then
        AddXMLField cmd, "flight_id", CStr(fInfo.flightID)
        ShowMessage "Resuming Flight " + CStr(fInfo.flightID), ACARSTEXTCOLOR
    End If
    
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent flight info", DEBUGTEXTCOLOR
End Sub

Public Sub RequestPilotInfo(pilotID As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "pilot"

    'Set the pilot ID in the flags
    Set flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild flags
    AddXMLField flags, "pilot_id", pilotID
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent pilot info request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestRunwayInfo(AirportCode As String, runwayCode As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "navaid"

    'Set the navaid info in the flags
    Set flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild flags
    AddXMLField flags, "id", AirportCode
    AddXMLField flags, "runway", runwayCode
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent Runway Info request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestNavaidInfo(navaidName As String, hdg As String, radioName As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "navaid"

    'Set the navaid info in the flags
    Set flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild flags
    AddXMLField flags, "id", navaidName
    AddXMLField flags, "radio", radioName
    AddXMLField flags, "hdg", hdg
    ReqStack.Queue cmd

    If config.ShowDebug Then ShowMessage "Sent Navigation Aid request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestCharts(AirportCode As String)
    Dim doc As New DOMDocument
    Dim cmd As IXMLDOMElement
    Dim flags As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "charts"

    'Set the navaid info in the flags
    Set flags = doc.createNode(NODE_ELEMENT, "flags", "")
    cmd.appendChild flags
    AddXMLField flags, "id", UCase(AirportCode)
    ReqStack.Queue cmd

    If config.ShowDebug Then ShowMessage "Sent Approach Chart request", DEBUGTEXTCOLOR
End Sub

Public Sub SendChat(msgText As String, Optional msgTo As String)
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("text")

    'Add the text and the recipient
    AddXMLField cmd, "text", msgText
    If (Len(msgTo) > 0) Then AddXMLField cmd, "to", msgTo
    ReqStack.Queue cmd
    
    'Show outgoing message
    ShowMessage "<" & frmMain.txtPilotID.Text & IIf(Len(msgTo) > 0, "->" & msgTo, "") & "> " & msgText, SELFCHATCOLOR
    If config.ShowDebug Then ShowMessage "Sent chat message", DEBUGTEXTCOLOR
End Sub

Public Function SendCredentials(userID As String, pwd As String) As Long
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("auth")

    'Add user and password
    AddXMLField cmd, "user", userID
    AddXMLField cmd, "password", pwd
    ReqStack.Queue cmd

    If config.ShowDebug Then ShowMessage "Logging In", DEBUGTEXTCOLOR
    SendCredentials = ReqStack.RequestID
End Function

Public Sub SendPing()
    Dim cmd As IXMLDOMElement

    'Build the request and send it
    ReqStack.Queue buildCMD("ping")
    If config.ShowDebug Then ShowMessage "Sent ping", DEBUGTEXTCOLOR
End Sub

Public Sub SendEndFlight()
    Dim cmd As IXMLDOMElement

    'Build the request and send it
    ReqStack.Queue buildCMD("end_flight")
    If config.ShowDebug Then ShowMessage "Sent end_flight message", DEBUGTEXTCOLOR
End Sub

Public Sub RequestEquipment()
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "eqList"
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent equipment list request", DEBUGTEXTCOLOR
End Sub

Public Sub RequestAirports()
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("datareq")
    AddXMLField cmd, "reqtype", "apList"
    ReqStack.Queue cmd
    
    If config.ShowDebug Then ShowMessage "Sent airport list request", DEBUGTEXTCOLOR
End Sub

Public Sub SendPosition(ByVal cPos As PositionData)
    Dim cmd As IXMLDOMElement
    Dim utcDate As Date
    Dim flags As Integer

    'Build the request
    Set cmd = buildCMD("position")
    utcDate = cPos.DateTime.UTCTime
    
    'Add position info nodes.
    AddXMLField cmd, "lat", cPos.Latitude
    AddXMLField cmd, "lon", cPos.Longitude
    AddXMLField cmd, "msl", cPos.AltitudeMSL
    AddXMLField cmd, "agl", cPos.AltitudeAGL
    AddXMLField cmd, "hdg", cPos.Heading
    AddXMLField cmd, "aSpeed", cPos.AirSpeed
    AddXMLField cmd, "gSpeed", cPos.GroundSpeed
    AddXMLField cmd, "vSpeed", cPos.VerticalSpeed
    AddXMLField cmd, "mach", Format(cPos.Mach, "0.000")
    AddXMLField cmd, "n1", Format(cPos.AverageN1, "##0.0")
    AddXMLField cmd, "n2", Format(cPos.AverageN2, "##0.0")
    AddXMLField cmd, "phase", cPos.Phase
    AddXMLField cmd, "simrate", CStr(cPos.simRate / 256)
    AddXMLField cmd, "flaps", CStr(cPos.Flaps)
    AddXMLField cmd, "fuel", CStr(cPos.Fuel)
    AddXMLField cmd, "weight", CStr(cPos.Weight)
    AddXMLField cmd, "date", Format(utcDate, "mm/dd/yyyy hh:nn:ss")

    'Build the flags
    If cPos.Paused Then flags = flags Or FLIGHTPAUSED
    If cPos.Slewing Then flags = flags Or FLIGHTSLEWING
    If cPos.Parked Then flags = flags Or FLIGHTPARKED
    If cPos.onGround Then flags = flags Or FLIGHTONGROUND
    If cPos.Spoilers Then flags = flags Or FLIGHT_SP_ARM
    If cPos.GearDown Then flags = flags Or FLIGHT_GEAR_DOWN
    If cPos.AP_NAV Then flags = flags Or FLIGHT_AP_NAV
    If cPos.AP_GPS Then flags = flags Or FLIGHT_AP_GPS
    If cPos.AP_HDG Then flags = flags Or FLIGHT_AP_HDG
    If cPos.AP_APR Then flags = flags Or FLIGHT_AP_APR
    If cPos.AP_ALT Then flags = flags Or FLIGHT_AP_ALT
    If cPos.AT_IAS Then flags = flags Or FLIGHT_AT_IAS
    If cPos.AT_MCH Then flags = flags Or FLIGHT_AT_MACH
    AddXMLField cmd, "flags", CStr(flags)

    'Send the request
    ReqStack.Queue cmd
    If config.ShowDebug Then ShowMessage "Sent position update", DEBUGTEXTCOLOR
End Sub

Public Function SendPIREP(info As FlightData) As Long
    Dim cmd As IXMLDOMElement

    'Build the request
    Set cmd = buildCMD("pirep")

    'Add pirep info fields
    AddXMLField cmd, "flightID", CStr(info.flightID)
    AddXMLField cmd, "flightcode", info.FlightNumber
    AddXMLField cmd, "eqType", info.EquipmentType
    AddXMLField cmd, "airportD", info.AirportD
    AddXMLField cmd, "airportA", info.AirportA
    AddXMLField cmd, "remarks", info.Remarks
    AddXMLField cmd, "network", info.Network
    AddXMLField cmd, "startTime", Format(info.StartTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "taxiOutTime", Format(info.TaxiOutTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "taxiFuel", CStr(info.TaxiFuel)
    AddXMLField cmd, "taxiWeight", CStr(info.TaxiWeight)
    AddXMLField cmd, "takeoffTime", Format(info.TakeoffTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "takeoffFuel", CStr(info.TakeoffFuel)
    AddXMLField cmd, "takeoffWeight", CStr(info.TakeoffWeight)
    AddXMLField cmd, "takeoffN1", Format(info.TakeoffN1, "##0.0")
    AddXMLField cmd, "takeoffSpeed", CStr(info.TakeoffSpeed)
    AddXMLField cmd, "landingTime", Format(info.LandingTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "landingFuel", CStr(info.LandingFuel)
    AddXMLField cmd, "landingWeight", CStr(info.LandingWeight)
    AddXMLField cmd, "landingN1", Format(info.LandingN1, "##0.0")
    AddXMLField cmd, "landingSpeed", CStr(info.LandingSpeed)
    AddXMLField cmd, "landingVSpeed", CStr(info.LandingVSpeed)
    AddXMLField cmd, "gateTime", Format(info.GateTime, "mm/dd/yyyy hh:nn:ss")
    AddXMLField cmd, "gateFuel", CStr(info.GateFuel)
    AddXMLField cmd, "gateWeight", CStr(info.GateWeight)

    'Send the request
    ReqStack.Queue cmd
    If config.ShowDebug Then ShowMessage "Sent Flight Report", DEBUGTEXTCOLOR
    SendPIREP = ReqStack.RequestID
End Function


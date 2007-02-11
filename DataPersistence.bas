Attribute VB_Name = "DataPersistence"
Option Explicit

Private Const MIN_BUILD = 63
Private Const HASH_SALT = "ha$h-Salt-ACARS-Value"

Private xdoc As DOMDocument
Private root As IXMLDOMElement

Public Sub PersistFlightData(ByVal writeMsg As Boolean, Optional fName As String = "ACARS Flight")
    Dim fNum As Integer
    Dim fRoot As String
    Dim xmlData As String, shaData As String
    Dim inf As IXMLDOMElement
    
    'Set critical error handler
    On Error GoTo FatalError
    frmMain.MousePointer = vbHourglass
    If config.ShowDebug Then ShowMessage "Generating XML Flight Data", DEBUGTEXTCOLOR
    
    'Create the DOM document if it doesn't exist
    If (xdoc Is Nothing) Then
        Dim SaveTime As New UTCDate
    
        SaveTime.SetNow
        Set xdoc = New DOMDocument
        Set root = xdoc.createNode(NODE_ELEMENT, "flight", "")
        root.setAttribute "created", I18nOutputDateTime(SaveTime)
        root.setAttribute "build", CStr(App.Revision)
        AddXMLField root, "aircraftAGL", CStr(acInfo.BaseAGL), False
        xdoc.appendChild root
    End If
    
    'Detach the inf node if it exists
    Set inf = root.selectSingleNode("info")
    If Not (inf Is Nothing) Then root.removeChild inf
    
    'Save the flight information
    Set inf = xdoc.createNode(NODE_ELEMENT, "info", "")
    AddXMLField inf, "id", CStr(info.FlightID), False
    AddXMLField inf, "airline", info.Airline.code, False
    AddXMLField inf, "flight", CStr(info.FlightNumber), False
    AddXMLField inf, "leg", CStr(info.FlightLeg), False
    AddXMLField inf, "phase", CStr(info.FlightPhase), False
    AddXMLField inf, "equipment", info.EquipmentType, False
    AddXMLField inf, "altitude", info.CruiseAltitude, False
    AddXMLField inf, "airportD", info.airportD.IATA, False
    AddXMLField inf, "airportDName", info.airportD.name, False
    AddXMLField inf, "airportA", info.AirportA.IATA, False
    AddXMLField inf, "airportAName", info.AirportA.name, False
    If Not (info.AirportL Is Nothing) Then
        AddXMLField inf, "airportL", info.AirportL.IATA, False
        AddXMLField inf, "airportLName", info.AirportL.name, False
    End If
        
    AddXMLField inf, "route", info.Route
    AddXMLField inf, "remarks", info.Remarks
    AddXMLField inf, "network", info.Network, False
    AddXMLField inf, "startTime", I18nOutputDateTime(info.StartTime)
    AddXMLField inf, "fs_ver", CStr(info.FSVersion), False
    If info.ScheduleVerified Then AddXMLField inf, "schedOK", "true", False
    If info.Offline Then AddXMLField inf, "offline", "true", False
    If (info.FlightPhase = COMPLETE) Then AddXMLField inf, "complete", "true", False
    If info.CheckRide Then AddXMLField inf, "checkRide", "true", False
    
    'Some of these fields may be zero but they should be persisted anyways
    AddXMLField inf, "taxiOutTime", I18nOutputDateTime(info.TaxiOutTime)
    AddXMLField inf, "taxiFuel", CStr(info.TaxiFuel), False
    AddXMLField inf, "taxiWeight", CStr(info.TaxiWeight), False
    AddXMLField inf, "takeoffTime", I18nOutputDateTime(info.TakeoffTime)
    AddXMLField inf, "takeoffFuel", CStr(info.TakeoffFuel), False
    AddXMLField inf, "takeoffWeight", CStr(info.TakeoffWeight), False
    AddXMLField inf, "takeoffN1", FormatNumber(info.TakeoffN1, "##0.0"), False
    AddXMLField inf, "takeoffSpeed", CStr(info.TakeoffSpeed), False
    AddXMLField inf, "landingTime", I18nOutputDateTime(info.LandingTime)
    AddXMLField inf, "landingFuel", CStr(info.LandingFuel), False
    AddXMLField inf, "landingWeight", CStr(info.LandingWeight), False
    AddXMLField inf, "landingN1", FormatNumber(info.LandingN1, "##0.0"), False
    AddXMLField inf, "landingSpeed", CStr(info.LandingSpeed), False
    AddXMLField inf, "landingVSpeed", CStr(info.LandingVSpeed), False
    AddXMLField inf, "gateTime", I18nOutputDateTime(info.GateTime)
    AddXMLField inf, "gateFuel", CStr(info.GateFuel), False
    AddXMLField inf, "gateWeight", CStr(info.GateWeight), False
    AddXMLField inf, "shutdownTime", I18nOutputDateTime(info.ShutdownTime)
    AddXMLField inf, "shutdownFuel", CStr(info.ShutdownFuel), False
    AddXMLField inf, "shutdownWeight", CStr(info.ShutdownWeight), False
    AddXMLField inf, "time0X", CStr(info.TimePaused), False
    AddXMLField inf, "time1X", CStr(info.TimeAt1X), False
    AddXMLField inf, "time2X", CStr(info.TimeAt2X), False
    AddXMLField inf, "time4X", CStr(info.TimeAt4X), False
    root.appendChild inf
    
    'Save the position cache
    If Positions.HasData Then
        Dim cache As IXMLDOMElement
        Dim Queue As Variant
        Dim x As Integer, startOfs As Integer
        Dim cPos As PositionData
        Dim e As IXMLDOMElement
        
        'Create the positions element if it doesn't exist
        Set cache = root.selectSingleNode("positions")
        If (cache Is Nothing) Then
            Set cache = xdoc.createNode(NODE_ELEMENT, "positions", "")
            root.appendChild cache
            startOfs = 0
        Else
            startOfs = CInt(cache.getAttribute("max"))
        End If
        
        'Save the offline position cache
        Queue = Positions.Queue
        For x = startOfs To UBound(Queue)
            Set e = xdoc.createNode(NODE_ELEMENT, "position", "")
            e.setAttribute "id", CStr(x + 1)
            Set cPos = Queue(x)
            
            'Write position data fields
            AddXMLField e, "lat", FormatNumber(cPos.Latitude, "#0.00000"), False
            AddXMLField e, "lon", FormatNumber(cPos.Longitude, "##0.00000"), False
            AddXMLField e, "msl", cPos.AltitudeMSL, False
            AddXMLField e, "agl", cPos.AltitudeAGL, False
            AddXMLField e, "hdg", cPos.Heading, False
            AddXMLField e, "aSpeed", cPos.Airspeed, False
            AddXMLField e, "gSpeed", cPos.GroundSpeed, False
            AddXMLField e, "vSpeed", cPos.VerticalSpeed, False
            AddXMLField e, "pitch", FormatNumber(cPos.Pitch, "#0.000"), False
            AddXMLField e, "bank", FormatNumber(cPos.Bank, "#0.000"), False
            AddXMLField e, "mach", FormatNumber(cPos.Mach, "0.000"), False
            AddXMLField e, "n1", FormatNumber(cPos.AverageN1, "##0.0"), False
            AddXMLField e, "n2", FormatNumber(cPos.AverageN2, "##0.0"), False
            AddXMLField e, "aoa", FormatNumber(cPos.AngleOfAttack, "#0.000"), False
            AddXMLField e, "g", FormatNumber(cPos.GForce, "#0.000"), False
            AddXMLField e, "fuelFlow", CStr(cPos.FuelFlow), False
            AddXMLField e, "phase", cPos.phase, False
            AddXMLField e, "simrate", CStr(cPos.simRate / 256), False
            AddXMLField e, "flaps", CStr(cPos.Flaps), False
            AddXMLField e, "fuel", CStr(cPos.Fuel), False
            AddXMLField e, "weight", CStr(cPos.weight), False
            AddXMLField e, "wHdg", CStr(cPos.WindHeading), False
            AddXMLField e, "wSpeed", CStr(cPos.WindSpeed), False
            AddXMLField e, "frameRate", CStr(cPos.FrameRate), False
            AddXMLField e, "date", I18nOutputDateTime(cPos.DateTime)
            AddXMLField e, "flags", CStr(cPos.Flags), False

            'Add to the positions
            cache.appendChild e
        Next
    
        'Save cache size
        cache.setAttribute "max", CStr(x)
    End If
    
    'Save the XML data
    Dim sha As New SHA256
    xmlData = "<?xml version=""1.0"" encoding=""UTF-8"" ?>" + vbCrLf + xdoc.XML
    shaData = sha.SHA256(xmlData, HASH_SALT)
    
    'Write the XML to a file
    fNum = FreeFile()
    fRoot = App.path + "\" + fName + " " + SavedFlightID(info)
    Open fRoot + ".xml" For Output As fNum
    Print #fNum, xmlData
    Close #fNum
    
    'Write the SHA256 hash to a file
    fNum = FreeFile()
    Open fRoot + ".sha" For Output As fNum
    Print #fNum, shaData
    Close #fNum
    
    'Mark us down as saved
    If writeMsg Then
        ShowMessage "Saved Flight Data", ACARSTEXTCOLOR
    ElseIf config.ShowDebug Then
        ShowMessage "Saved Flight Data (" + CStr(Positions.Size) + " position records)", DEBUGTEXTCOLOR
    End If
    
ExitSub:
    frmMain.MousePointer = vbDefault
    Exit Sub
    
FatalError:
    MsgBox "Error saving Flight Data: " & err.Description, vbCritical, "Error"
    Resume ExitSub
    
End Sub

Private Function hasFlag(ByVal f As Long, ByVal attr As Long) As Boolean
    hasFlag = ((f And attr) <> 0)
End Function

Public Function RestoreFlightData(ByVal flightCode As String) As SavedFlight
    Dim fNum As Integer
    Dim fRoot As String, shaHash As String, shaCalc As String
    Dim doc As New DOMDocument, root As IXMLDOMElement
    Dim xmlData As String
    Dim result As New SavedFlight
    Dim sha As New SHA256
    
    'Set critical error handler
    On Error GoTo FatalError

    'Calculate the file name root
    frmMain.MousePointer = vbHourglass
    fRoot = App.path + "\ACARS Flight " + flightCode
    If Not FileExists(fRoot + ".sha") Or Not FileExists(fRoot + ".xml") Then Exit Function

    'Get the SHA256 hash value
    fNum = FreeFile()
    Open fRoot + ".sha" For Input As fNum
    shaHash = Input(LOF(fNum) - 2, fNum)
    Close #fNum
    
    'Load the flight data
    fNum = FreeFile()
    Open fRoot + ".xml" For Input As fNum
    xmlData = Input(LOF(fNum) - 2, fNum)
    Close #fNum
    
    'Validate the SHA256 hash
    shaCalc = sha.SHA256(xmlData, HASH_SALT)
    If (shaHash <> shaCalc) Then
        MsgBox "The Flight Data appears to have been altered!" + vbCrLf + vbCrLf + _
            "Validation Code = " + shaHash + vbCrLf + "Calculated Code = " + shaCalc, _
            vbCritical + vbOKOnly, "Altered Flight Data"
        GoTo ExitSub
    End If
    
    'Convert into XML document
    doc.async = False
    If Not doc.loadXML(xmlData) Then
        Dim strError As String
        Dim xmlError As IXMLDOMParseError
    
        Set xmlError = doc.parseError
        strError = "Error code: " & xmlError.errorCode & vbCrLf _
            & "Reason: " & xmlError.reason & vbCrLf _
            & "Source: " & vbCrLf & xmlError.srcText
        MsgBox "The following error occurred while parsing flight data: " & _
            vbCrLf & vbCrLf & strError, vbCritical Or vbOKOnly, "Fatal Error!"
        GoTo ExitSub
    End If
    
    'Get the root element
    Set root = doc.selectSingleNode("flight")
    If (root Is Nothing) Then
        ShowMessage "Invalid Saved Flight - No flight element", ACARSERRORCOLOR
        GoTo ExitSub
    End If
    
    'Validate the build number
    Dim buildNumber As Integer
    buildNumber = CInt(getAttr(root, "build", "63"))
    If (buildNumber < MIN_BUILD) Then
        MsgBox "This Flight was saved using Build " & CStr(buildNumber) & ". This version of" & _
            vbCrLf & "the ACARS client can only read Flights saved using Build " & CStr(MIN_BUILD) & _
            " or newer.", vbCritical + vbOKOnly, "Saved Flight Error"
        GoTo ExitSub
    ElseIf (buildNumber > App.Revision) Then
        MsgBox "This Flight was saved using a newer revision (Build " & CStr(buildNumber) & ")" & _
            vbCrLf & "of ACARS. Some saved Flight data may not be loaded.", vbExclamation + vbOKOnly, _
            "Newer ACARS Version"
    End If
    
    'Get the flight information node
    Dim inf As IXMLDOMElement
    Set inf = root.selectSingleNode("info")
    If (inf Is Nothing) Then
        ShowMessage "Invalid Saved Flight " + flightCode + " - No Flight Information", ACARSERRORCOLOR
        GoTo ExitSub
    End If
    
    'Load the flight information
    If config.ShowDebug Then ShowMessage "Restoring Flight Information", DEBUGTEXTCOLOR
    Set result.FlightInfo = New FlightData
    result.AircraftAGL = CInt(getChild(root, "aircrafAGL", "5"))
    With result.FlightInfo
        .FlightID = CLng(getChild(inf, "id", "0"))
        Set .Airline = config.GetAirline(getChild(inf, "airline", ""))
        .FlightNumber = CInt(getChild(inf, "flight", "1"))
        .FlightLeg = CInt(getChild(inf, "leg", "1"))
        .FlightPhase = CInt(getChild(inf, "phase", CStr(AIRBORNE)))
        .EquipmentType = getChild(inf, "equipment", "CRJ-200")
        .CruiseAltitude = getChild(inf, "altitude", "3000")
        .Network = getChild(inf, "network", "Offline")
        .CheckRide = CBool(getChild(inf, "checkRide", "false"))
        .ScheduleVerified = CBool(getChild(inf, "schedOK", "false"))
        Set .airportD = config.GetAirport(getChild(inf, "airportD", ""))
        Set .AirportA = config.GetAirport(getChild(inf, "airportA", ""))
        Set .AirportL = config.GetAirport(getChild(inf, "airportL", ""))
        .Route = getChild(inf, "route", "")
        .Remarks = getChild(inf, "remarks", "")
        .FSVersion = CInt(getChild(inf, "fs_ver", "7"))
        .StartTime.ParseUTCTime = I18nDateTime(getChild(inf, "startTime", ""))
        .TaxiOutTime.ParseUTCTime = I18nDateTime(getChild(inf, "taxiOutTime", ""))
        .TaxiFuel = CLng(getChild(inf, "taxiFuel", "0"))
        .TaxiWeight = CLng(getChild(inf, "taxiWeight", "0"))
        .TakeoffTime.ParseUTCTime = I18nDateTime(getChild(inf, "takeoffTime", ""))
        .TakeoffFuel = CLng(getChild(inf, "takeoffFuel", "0"))
        .TakeoffWeight = CLng(getChild(inf, "takeoffWeight", "0"))
        .TakeoffN1 = CDbl(getChild(inf, "takeoffN1", "00.0"))
        .TakeoffSpeed = CInt(getChild(inf, "takeoffSpeed", "0"))
        .LandingTime.ParseUTCTime = I18nDateTime(getChild(inf, "landingTime", ""))
        .LandingFuel = CLng(getChild(inf, "landingFuel", "0"))
        .LandingWeight = CLng(getChild(inf, "landingWeight", "0"))
        .LandingN1 = CDbl(getChild(inf, "landingN1", "00.0"))
        .LandingSpeed = CInt(getChild(inf, "landingSpeed", "0"))
        .LandingVSpeed = CInt(getChild(inf, "landingVSpeed", "0"))
        .GateTime.ParseUTCTime = I18nDateTime(getChild(inf, "gateTime", ""))
        If (Year(.GateTime.UTCTime) < 2000) Then .GateTime.Clone (.LandingTime)
        .GateFuel = CLng(getChild(inf, "gateFuel", "0"))
        .GateWeight = CLng(getChild(inf, "gateWeight", "0"))
        .ShutdownTime.ParseUTCTime = I18nDateTime(getChild(inf, "shutdownTime", ""))
        If (Year(.ShutdownTime.UTCTime) < 2000) Then .ShutdownTime.Clone (.LandingTime)
        .ShutdownFuel = CLng(getChild(inf, "shutdownFuel", "0"))
        .ShutdownWeight = CLng(getChild(inf, "shutdownWeight", "0"))
        .TimePaused = CLng(getChild(inf, "time0X", "0"))
        .TimeAt1X = CLng(getChild(inf, "time1X", "0"))
        .TimeAt2X = CLng(getChild(inf, "time2X", "0"))
        .TimeAt4X = CLng(getChild(inf, "time4X", "0"))
        If (.FlightPhase >= TAXI_IN) And (.FlightPhase <= COMPLETE) Then
            .FlightData = True
            .FlightPhase = COMPLETE
            .InFlight = False
        Else
            .InFlight = True
        End If
    End With
    
    'Load the position cache
    Dim pcache As IXMLDOMElement
    Set pcache = root.selectSingleNode("positions")
    If Not (pcache Is Nothing) Then
        Dim pList As IXMLDOMNodeList
        Dim p As IXMLDOMElement
        Dim sPos As PositionData
        Dim Flags As Long
    
        If config.ShowDebug Then ShowMessage "Restoring Offline Positions", DEBUGTEXTCOLOR
        Set pList = pcache.selectNodes("position")
        For Each p In pList
            Set sPos = New PositionData
            sPos.Latitude = ParseNumber(getChild(p, "lat", "0.00"))
            sPos.Longitude = ParseNumber(getChild(p, "lon", "0.00"))
            sPos.AltitudeMSL = CLng(getChild(p, "msl", "0"))
            sPos.AltitudeAGL = CLng(getChild(p, "agl", "0"))
            sPos.Heading = CInt(getChild(p, "hdg", "0"))
            sPos.Airspeed = CInt(getChild(p, "aSpeed", "0"))
            sPos.GroundSpeed = CInt(getChild(p, "gSpeed", "0"))
            sPos.VerticalSpeed = CInt(getChild(p, "vSpeed", "0"))
            sPos.Pitch = ParseNumber(getChild(p, "pitch", "0.00"))
            sPos.Bank = ParseNumber(getChild(p, "bank", "0.00"))
            sPos.Mach = ParseNumber(getChild(p, "mach", "0.00"))
            sPos.AngleOfAttack = ParseNumber(getChild(p, "aoa", "0.00"))
            sPos.GForce = ParseNumber(getChild(p, "g", "0.00"))
            
            'This is a hack since we don't have individual N1/N2 values
            sPos.EngineCount = 1
            sPos.setN1 0, ParseNumber(getChild(p, "n1", "00.00"))
            sPos.setN2 0, ParseNumber(getChild(p, "n2", "00.00"))
            sPos.setFuelFlow 0, ParseNumber(getChild(p, "fuelFlow", "0"))
            
            sPos.phase = getChild(p, "phase", "Airborne")
            sPos.simRate = CInt(getChild(p, "simrate", "256") / 256)
            sPos.Flaps = CInt(getChild(p, "flaps", "0"))
            sPos.Fuel = CLng(getChild(p, "fuel", "0"))
            sPos.weight = CLng(getChild(p, "weight", "0"))
            sPos.WindHeading = CInt(getChild(p, "wHdg", "0"))
            sPos.WindSpeed = CInt(getChild(p, "wSpeed", "0"))
            sPos.FrameRate = CInt(getChild(p, "frameRate", "0"))
            sPos.DateTime.ParseUTCTime = I18nDateTime(getChild(p, "date", CStr(Now)))
            
            'Build the flags
            Flags = CLng(getChild(p, "flags", "0"))
            sPos.Paused = hasFlag(Flags, FLIGHT_PAUSED)
            sPos.Touchdown = hasFlag(Flags, FLIGHT_TOUCHDOWN)
            sPos.Parked = hasFlag(Flags, FLIGHT_PARKED)
            sPos.onGround = hasFlag(Flags, FLIGHT_ONGROUND)
            sPos.Spoilers = hasFlag(Flags, FLIGHT_SP_ARM)
            sPos.GearDown = hasFlag(Flags, FLIGHT_GEAR_DOWN)
            sPos.AfterBurner = hasFlag(Flags, FLIGHT_AFTERBURNER)
            sPos.Overspeed = hasFlag(Flags, FLIGHT_OVERSPEED)
            sPos.Stall = hasFlag(Flags, FLIGHT_STALL)
            sPos.AP_NAV = hasFlag(Flags, FLIGHT_AP_NAV)
            sPos.AP_GPS = hasFlag(Flags, FLIGHT_AP_GPS)
            sPos.AP_HDG = hasFlag(Flags, FLIGHT_AP_HDG)
            sPos.AP_APR = hasFlag(Flags, FLIGHT_AP_APR)
            sPos.AP_ALT = hasFlag(Flags, FLIGHT_AP_ALT)
            sPos.AT_IAS = hasFlag(Flags, FLIGHT_AT_IAS)
            sPos.AT_MCH = hasFlag(Flags, FLIGHT_AT_MACH)
            
            'Add to cache
            result.AddPosition sPos
        Next
    End If
    
    'Save the XML
    result.XML = xmlData
    Set RestoreFlightData = result

ExitSub:
    frmMain.MousePointer = vbDefault
    Exit Function
    
FatalError:
    MsgBox "Error loading Flight Data: " & err.Description, vbCritical, "Error"
    Resume ExitSub

End Function

Public Function SavedFlights() As Variant
    Dim results As Variant
    Dim fName As String
    Dim isOK As Boolean
    
    results = Array()
    fName = Dir(App.path + "\ACARS Flight *.sha", vbNormal)
    While (fName <> "")
        fName = Left(fName, InStr(0, fName, ".") - 1) 'Get the base name
        isOK = FileExists(App.path + "\" + fName + ".xml")
        isOK = isOK And FileExists(config.FS9Files + "\" + fName + ".FLT")
        isOK = isOK And FileExists(config.FS9Files + "\" + fName + ".WX")
        If isOK Then
            ReDim Preserve results(UBound(results) + 1)
            results(UBound(results)) = fName
        End If
        
        fName = Dir$()
    Wend
    
    SavedFlights = results
End Function

Public Sub DeleteSavedFlight(ByVal flightCode As String, Optional KillAll As Boolean = True)
    Dim fRoot As String
    
    On Error Resume Next
    fRoot = App.path + "\ACARS Flight " + flightCode
    
    'Kill the saved ACARS data
    If KillAll Then
        Kill fRoot + ".xml"
        Kill fRoot + ".sha"
    End If
    
    'Delete the saved flight
    If Not config.WideFSInstalled Then
        Kill config.FS9Files + "\" + "ACARS Flight " + flightCode + ".FLT"
        Kill config.FS9Files + "\" + "ACARS Flight " + flightCode + ".WX"
    End If
    
    'Display message
    If config.ShowDebug Then ShowMessage "Deleted persisted data for Flight " + _
        flightCode, DEBUGTEXTCOLOR
End Sub

Public Function FileExists(ByVal fPath As String) As Boolean
    FileExists = (Dir(fPath) <> "")
End Function

Public Function SavedFlightID(fInfo As FlightData) As String
    If (fInfo.FlightID > 0) Then
        SavedFlightID = Format(fInfo.FlightID, "00000")
    Else
        SavedFlightID = "O-" + Format(info.StartTime.UTCTime, "yyyymmddhh")
    End If
End Function

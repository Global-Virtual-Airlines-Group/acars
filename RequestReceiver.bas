Attribute VB_Name = "RequestReceiver"
Option Explicit

Public ACKStack As New ACKBuffer
Public Const XMLRESPONSEROOT = "ACARSResponse"

Public Function WaitForACK(ByVal ID As Long, Optional ByVal TimeOut As Long = 2500) As Boolean
    Dim gotACK As Boolean
    Dim totalTime As Long
    Dim pingTime As Integer
    
    ACKStack.Queue ID
    
    'If we're waiting to send something, then send it
    If ReqStack.HasData Then ReqStack.Send
    While (totalTime < TimeOut) And Not gotACK
        gotACK = ACKStack.HasReceived(ID)
        If Not gotACK Then
            totalTime = totalTime + 100
            pingTime = pingTime + 100
            If (pingTime >= 1000) Then
                If config.ShowDebug Then ShowMessage "Pinging Socket", DEBUGTEXTCOLOR
                SendPing
                ReqStack.Send
                pingTime = 0
            End If
            
            Sleep 100
            DoEvents
        End If
    Wend
    
    WaitForACK = gotACK
End Function

Public Sub ProcessMessage(ByVal msgText As String)
    Dim doc As New DOMDocument
    Dim root As IXMLDOMElement
    Dim cmds As IXMLDOMNodeList
    Dim cmd As IXMLDOMElement
    Dim err As IXMLDOMElement
    
    'Load the XML
    doc.async = False
    If Not doc.loadXML(msgText) Then
        Dim strError As String
        Dim xmlError As IXMLDOMParseError
        
        Set xmlError = doc.parseError
        strError = "Error code: " & xmlError.errorCode & vbCrLf _
            & "Reason: " & xmlError.reason & vbCrLf _
            & "Source: " & vbCrLf & xmlError.srcText
        MsgBox "The following fatal error occurred while parsing XML from the server:" & vbCrLf & vbCrLf & strError, vbCritical Or vbOKOnly, "Fatal Error!"
        If config.ShowDebug Then ShowDebug msgText, DEBUGTEXTCOLOR
        Exit Sub
    Else
        If config.ShowDebug Then ShowDebug msgText, XML_IN_COLOR
    End If
    
    'Grab all CMD nodes
    Set root = doc.documentElement
    Set cmds = root.selectNodes("CMD")

    'Process each CMD response. First check for an error element
    For Each cmd In cmds
        Set err = cmd.selectSingleNode("error")
        If (err Is Nothing) Then
            ProcessResponse cmd
        Else
            Dim ReqID As Long
        
            'Take appropriate action based on request ID specified in the error.
            ReqID = Val("&H" + getAttr(cmd, "id", "0"))
            PlaySoundFile "notify_error.wav"
            Select Case ReqID
                Case info.AuthReqID
                    info.AuthReqID = 0
                    frmMain.CloseACARSConnection
                    MsgBox "The following error occurred while attempting to connect to the ACARS server:" & vbCrLf & vbCrLf & getChild(cmd, "error", "?"), vbOKOnly Or vbExclamation, "Error"
                Case Else
                    MsgBox "The following error occurred:" & vbCrLf & vbCrLf & getChild(cmd, "error", "?"), vbOKOnly Or vbExclamation, "Error"
            End Select
        End If
    Next
End Sub

Private Sub ProcessResponse(cmd As IXMLDOMNode)
    Dim cmdName As String
    
    'Set critical error handler
    On Error GoTo ErrorHandler
    
    'Branch depending on type of command.
    cmdName = LCase(getAttr(cmd, "type", ""))
    Select Case cmdName
        Case "ack"
            ProcessACK cmd
        Case "datarsp"
            ProcessDataResponse cmd
        Case "text"
            ProcessChatText cmd
        Case "smsg"
            ProcessServerMsg cmd
        Case Else
            ShowMessage "Unknown command - " + cmdName, ACARSERRORCOLOR
    End Select
    
ExitSub:
    Exit Sub
    
ErrorHandler:
    ShowMessage "Error Processing ACARS " + cmdName + " Response - " + err.Description, ACARSERRORCOLOR
    Resume ExitSub
End Sub

Private Sub ProcessACK(cmdNode As IXMLDOMNode)
    Dim ReqID As Long
    Dim RequestInfo As Boolean
    
    ReqID = Val("&H" + getAttr(cmdNode, "id", "0"))
    If ACKStack.HasReceived(ReqID) Then Exit Sub
    
    If config.ShowDebug Then ShowMessage "Received ACK for " + Hex(ReqID), DEBUGTEXTCOLOR
    ACKStack.Receive ReqID
    
    'Check if we are requesting flight information
    RequestInfo = CBool(getChild(cmdNode, "sendInfo", "false"))
    If RequestInfo Then
        ShowMessage "Requesting Flight Information", ACARSTEXTCOLOR
        info.InfoReqID = SendFlightInfo(info)
        ReqStack.Send
    End If
    
    'If the ACK ID is zero, abort
    If (ReqID = 0) Then Exit Sub
            
    'Do special things if we're responding to a message
    If ReqID = info.AuthReqID Then
        Dim newBuild As Integer
        Dim RoleNames As Variant
        Dim rName As Variant
    
        'Enable stuff and update status
        frmMain.sbMain.Panels(1).Text = "Status: Logged in to ACARS server"
        frmMain.SSTab1.TabEnabled(1) = True
        frmMain.mnuOptionsFlyOffline.Enabled = False
        
        'Display newer build if available
        newBuild = CInt(getChild(cmdNode, "latestBuild", "0"))
        If (newBuild > App.Revision) Then
            ShowMessage "A new ACARS version (Build " + CStr(newBuild) + ") is available.", ACARSTEXTCOLOR
            PlaySoundFile "notify_newversion.wav"
        Else
            PlaySoundFile "notify_welcome.wav"
        End If
        
        'Get role names
        config.ClearRoles
        RoleNames = Split(getChild(cmdNode, "roles", ""), ",")
        For Each rName In RoleNames
            config.AddRole rName
        Next
        
        'Get equipment ratings
        RoleNames = Split(getChild(cmdNode, "ratings", ""), ",")
        For Each rName In RoleNames
            config.AddRating rName
        Next
        
        'Check our acces with FS not running
        Dim NeedsFS As Boolean
        NeedsFS = Not (config.HasRole("HR") Or config.HasRole("PIREP") Or config.HasRole("Dispatch"))
        If NeedsFS And Not IsFSRunning Then
            frmMain.ToggleACARSConnection True
            MsgBox "Microsoft Flight Simulator must be started before you connect to the ACARS Server.", _
                vbOKOnly + vbExclamation, "Flight Simulator not started"
            Exit Sub
        End If
        
        'Hide the stealth box
        If Not config.HasRole("HR") Then
            frmMain.chkStealth.value = 0
            frmMain.chkStealth.visible = False
            frmMain.chkStealth.Enabled = False
        End If
        
        'Get unrestricted flag
        config.NoMessages = CBool(getChild(cmdNode, "noMsgs", "false"))
        config.IsUnrestricted = CBool(getChild(cmdNode, "unrestricted", "false")) And Not config.NoMessages
        If config.IsUnrestricted And config.ShowDebug Then
            ShowMessage "Unrestricted messaging operations", DEBUGTEXTCOLOR
        ElseIf config.NoMessages Then
            ShowMessage "ACARS Messaging disabled", SYSMSGCOLOR
        End If
        
        'Check our time offset
        Dim timeOfs As Long
        timeOfs = CLng(getChild(cmdNode, "timeOffset", "0"))
        If (Abs(timeOfs) > 900) Then ShowMessage "Your System Clock is off by " & CStr(timeOfs) & _
            " seconds!", ACARSERRORCOLOR
        
        'Update Equipment/Airport comboboxes
        If Not config.IsConfigUpToDate Then
            ShowMessage "Updating Airport/Airline/Equipment lists", ACARSTEXTCOLOR
            RequestEquipment
            RequestAirports
            RequestAirlines
        End If
        
        'Tell the gauge we are connected
        GAUGE_SetStatus ACARS_CONNECTED
        
        'Request Pilot List/SB3 Privat voice
        If config.SB3Support Then RequestPrivateVoiceURL
        RequestPilotList
        
        'Get a flight ID if we have a flight
        If (info.InFlight Or info.FlightData) And Not (RequestInfo) Then info.InfoReqID = SendFlightInfo(info)
        
        'Send data
        ReqStack.Send
        DoEvents
    ElseIf (ReqID = info.InfoReqID) Then
        Dim newID As Long
        Dim isCheckRide As Boolean
    
        'Get the flight ID
        newID = CLng(getChild(cmdNode, "flight_id", "0"))
        If (info.FlightID = 0) Then
            ShowMessage "Assigned Flight ID " + CStr(newID), ACARSTEXTCOLOR
        ElseIf (newID <> info.FlightID) Then
            ShowMessage "Requested Flight ID " + CStr(info.FlightID) + ", received " + CStr(newID), ACARSTEXTCOLOR
        End If
        
        'Check the check ride flag
        isCheckRide = CBool(getChild(cmdNode, "checkRide", "false"))
        If isCheckRide Then
            isCheckRide = (MsgBox("You have a pending " & info.EquipmentType & " Check Ride." & _
                vbCrLf & vbCrLf & "Are you flying the Check Ride now?", vbQuestion + vbYesNo, _
                "Pending Check Ride") = vbYes)
        End If
        
        'Delete any old persisted flight and then save the flight with the new ID
        DeleteSavedFlight SavedFlightID(info)

        'Save the flight ID to the registry for crash recovery.
        info.FlightID = newID
        If isCheckRide Then
            info.CheckRide = True
            frmMain.chkCheckRide.value = 1
            ShowMessage "ACARS Check Ride detected", SYSMSGCOLOR
        ElseIf info.CheckRide And Not isCheckRide Then
            info.CheckRide = False
            frmMain.chkCheckRide.value = 0
            MsgBox "You do not currently have a pending Check Ride in the " & info.EquipmentType & _
                ".", vbExclamation + vbOKOnly, "No Pending Check Ride"
        End If
            
        config.SaveFlightCode SavedFlightID(info)
        PersistFlightData True
        SaveFlight
    End If
End Sub

Private Sub ProcessServerMsg(cmdNode As IXMLDOMNode)
    Dim msgs As IXMLDOMNodeList
    Dim Msg As IXMLDOMNode

    'Branch depending on type.
    Select Case LCase(getAttr(cmdNode, "msgtype", "text"))
        Case "text"
            Set msgs = cmdNode.selectNodes("text")
            For Each Msg In msgs
                ShowMessage Msg.Text, SYSMSGCOLOR
            Next
    End Select
End Sub

Private Sub ProcessChatText(cmdNode As IXMLDOMNode)
    Dim msgFrom As String
    Dim msgTo As String
    Dim msgText As String
    Dim p As Pilot
    Dim IsSterile As Boolean
    
    'Check if we are in sterile cockpit mode
    IsSterile = False
    If ((info.FlightPhase = AIRBORNE) Or (info.FlightPhase = TAKEOFF) Or (info.FlightPhase = ROLLOUT)) Then
        If Not (pos Is Nothing) Then IsSterile = (pos.AltitudeMSL < 10000)
    End If
    
    'Get the message info
    msgTo = getChild(cmdNode, "to", "")
    msgText = getChild(cmdNode, "text", "")
    msgFrom = getChild(cmdNode, "from", "SYSTEM")
    If (msgFrom <> "SYSTEM") Then Set p = users.GetPilot(msgFrom)
    If (Not (p Is Nothing) And config.ShowPilotNames) Then msgFrom = p.name

    'Check if there is a "To" element. If so, it's a private message.
    config.MsgReceived = True
    If (msgTo = "") Then
        If (config.HideMessagesWhenBusy And config.Busy) Then
            If config.ShowDebug Then ShowMessage "Ignoring chat message - Busy", DEBUGTEXTCOLOR
        ElseIf (config.SterileCockpit And IsSterile) Then
            If config.ShowDebug Then ShowMessage "Ignoring chat message - Sterile Cockpit", DEBUGTEXTCOLOR
        Else
            ShowMessage "<" & msgFrom & "> " & msgText, PUBLICCHATCOLOR
            GAUGE_SetChat False
            
            'Play a sound if the option is on
            If (config.PlaySound And (Not config.Busy) And (GetForegroundWindow <> frmMain.hWnd)) _
                Then PlaySoundFile "notify_msg.wav"
        End If
    Else
        ShowMessage "<" & msgFrom & ":PRIVATE> " & msgText, PRIVATECHATCOLOR
        GAUGE_SetChat True
        
        'Send a response if we are busy
        If Not (p Is Nothing) And (Left(msgText, 5) <> "AUTO:") Then
            If config.Busy Then
                If (config.BusyMessage <> "") Then
                    SendChat "AUTO: " & config.BusyMessage, msgFrom
                Else
                    SendChat "AUTO: I am currently busy and not available to chat.", msgFrom
                End If
                
                ReqStack.Send
            ElseIf (config.SterileCockpit And IsSterile) Then
                SendChat "AUTO: I am in a Sterile Cockpit environment and not available to chat.", msgFrom
                ReqStack.Send
            ElseIf (config.PlaySound And (Not config.Busy) And (GetForegroundWindow <> frmMain.hWnd)) Then
                PlaySoundFile "notify_msg.wav"
            End If
        End If
    End If
End Sub

Private Sub ProcessDataResponse(cmdNode As IXMLDOMNode)
    Dim rspNodes As IXMLDOMNodeList
    Dim rspNode As IXMLDOMNode
    
    Dim pNodes As IXMLDOMNodeList
    Dim pNode As IXMLDOMNode
    Dim ID As String, ReqID As Long
    Dim freq As String
    Dim Msg As String
    
    Dim p As Pilot
    Dim OldDispatch As Boolean, NewDispatch As Boolean
    
    'Check if this is a response to a specific request
    ReqID = Val("&H" + getAttr(cmdNode, "id", "0"))
    If (ReqID <> 0) Then
        ACKStack.Receive ReqID
        If config.ShowDebug Then ShowMessage "Received ACK for " + Hex(ReqID), DEBUGTEXTCOLOR
        
        'Check if this is the schedule search request flag
        If (ReqID = info.SchedReqID) Then
            Set pNode = cmdNode.selectSingleNode("flights")
            info.ScheduleVerified = (Not (pNode Is Nothing)) And (pNode.childNodes.Length > 0)
            info.SchedReqID = 0
            Exit Sub
        End If
    End If
    
    'Branch based on response type
    Set rspNodes = cmdNode.selectNodes("rsptype")
    For Each rspNode In rspNodes
        Select Case LCase(rspNode.Text)
            Case "pilotlist"
                Dim name As String
                OldDispatch = users.DispatchOnline
                
                If config.ShowDebug Then ShowMessage "Updating Pilot List", DEBUGTEXTCOLOR
                users.ClearPilots
                Set pNodes = cmdNode.selectSingleNode("pilotlist").selectNodes("Pilot")
                For Each pNode In pNodes
                    Set p = New Pilot
                    p.ID = getAttr(pNode, "id", "")
                    p.FirstName = getChild(pNode, "firstname", "")
                    p.LastName = getChild(pNode, "lastname", "")
                    p.EquipmentType = getChild(pNode, "eqtype", "CRJ-200")
                    p.Rank = getChild(pNode, "rank", "First Officer")
                    p.Legs = CInt(Replace(getChild(pNode, "legs", "0"), ".", config.DecimalSeparator))
                    p.Hours = CDbl(Replace(getChild(pNode, "hours", "0"), ".", config.DecimalSeparator))
                    p.flightCode = getChild(pNode, "flightCode", "")
                    p.FlightEQ = getChild(pNode, "flightEQ", "")
                    Set p.airportD = config.GetAirport(getChild(pNode, "airportD", ""))
                    Set p.AirportA = config.GetAirport(getChild(pNode, "airportA", ""))
                    p.ClientBuild = CInt(getChild(pNode, "clientBuild", "60"))
                    p.RemoteAddress = getChild(pNode, "remoteaddr", "???")
                    p.RemoteHost = getChild(pNode, "remotehost", "???")
                    p.IsBusy = CBool(getChild(pNode, "isBusy", "false"))
                    p.IsHidden = CBool(getChild(pNode, "isHidden", "false"))
                    p.SetRoles getChild(pNode, "roles", "Pilot")
                
                    'Check if the ID is ours
                    If (p.ID <> "") Then
                        users.AddPilot p
                        If (UCase(p.ID) = UCase(frmMain.txtPilotID.Text)) Then
                            frmMain.txtPilotName.Text = p.name
                            frmMain.txtPilotName.visible = True
                            frmMain.lblName.visible = True
                        End If
                    End If
                Next
                
                'Update the pilot list
                users.UpdatePilotList
                
                'If HR/Dispatch is now online, log it
                NewDispatch = users.DispatchOnline
                If (NewDispatch And (OldDispatch <> NewDispatch)) Then
                    If (config.HasRole("HR") Or config.HasRole("Dispatch")) Then
                        ShowMessage "ACARS Messaging Restrictions waived by your login", ACARSTEXTCOLOR
                    ElseIf Not (info.InFlight Or config.IsUnrestricted) Then
                        ShowMessage "ACARS Messaging Restrictions waived", ACARSTEXTCOLOR
                    End If
                End If
                
            Case "addpilots"
                Dim oldPilot As Pilot
                OldDispatch = users.DispatchOnline
            
                'Process added pilots
                Set pNodes = cmdNode.selectSingleNode("addpilots").selectNodes("Pilot")
                For Each pNode In pNodes
                    Set p = New Pilot
                    p.ID = getAttr(pNode, "id", "")
                    p.FirstName = getChild(pNode, "firstname", "")
                    p.LastName = getChild(pNode, "lastname", "")
                    p.EquipmentType = getChild(pNode, "eqtype", "CRJ-200")
                    p.Rank = getChild(pNode, "rank", "First Officer")
                    p.Legs = CInt(Replace(getChild(pNode, "legs", "0"), ".", config.DecimalSeparator))
                    p.Hours = CDbl(Replace(getChild(pNode, "hours", "0"), ".", config.DecimalSeparator))
                    p.flightCode = getChild(pNode, "flightCode", "")
                    p.FlightEQ = getChild(pNode, "flightEQ", "")
                    Set p.airportD = config.GetAirport(getChild(pNode, "airportD", ""))
                    Set p.AirportA = config.GetAirport(getChild(pNode, "airportA", ""))
                    p.ClientBuild = CInt(getChild(pNode, "clientBuild", "60"))
                    p.RemoteAddress = getChild(pNode, "remoteaddr", "???")
                    p.RemoteHost = getChild(pNode, "remotehost", "???")
                    p.IsBusy = CBool(getChild(pNode, "isBusy", "false"))
                    p.IsHidden = CBool(getChild(pNode, "isHidden", "false"))
                    p.SetRoles getChild(pNode, "roles", "Pilot")

                    'Send login message
                    If (p.ID <> "") Then
                        Set oldPilot = users.GetPilot(p.ID)
                        users.AddPilot p
                        If (oldPilot Is Nothing) Then ShowMessage p.name + " (" + p.ID + _
                            ") logged into the ACARS server.", SYSMSGCOLOR
                    End If
                Next
                
                'Update the pilot list
                users.UpdatePilotList
                
                'If HR/Dispatch is now online, log it
                NewDispatch = users.DispatchOnline
                If (NewDispatch And (OldDispatch <> NewDispatch)) Then
                    If (config.HasRole("HR") Or config.HasRole("Dispatch")) Then
                        ShowMessage "ACARS Messaging Restrictions waived by your login", ACARSTEXTCOLOR
                    ElseIf Not (info.InFlight Or config.IsUnrestricted) Then
                        ShowMessage "ACARS Messaging Restrictions waived", ACARSTEXTCOLOR
                    End If
                End If
                
            Case "delpilots"
                OldDispatch = users.DispatchOnline
                
                'Process the pilot list
                Set pNodes = cmdNode.selectSingleNode("delpilots").selectNodes("Pilot")
                For Each pNode In pNodes
                    Set p = New Pilot
                    p.ID = getAttr(pNode, "id", "")
                    p.FirstName = getChild(pNode, "firstname", "")
                    p.LastName = getChild(pNode, "lastname", "")

                    users.DeletePilot p.ID
                    ShowMessage p.name + " (" + p.ID + ") logged out from the ACARS server.", SYSMSGCOLOR
                Next
                
                'Update the Pilot List
                users.UpdatePilotList
                
                'If HR/Dispatch is now offline, log it
                NewDispatch = users.DispatchOnline
                If (Not NewDispatch And (OldDispatch <> NewDispatch)) Then
                    If Not (info.InFlight Or config.IsUnrestricted) Then
                        ShowMessage "ACARS Messaging Restrictions restored", ACARSTEXTCOLOR
                    End If
                End If
                
            Case "atc"
                Dim ctr As Controller
                
                If config.ShowDebug Then ShowMessage "Updating ATC List", DEBUGTEXTCOLOR
                Set pNodes = cmdNode.selectSingleNode("atc").selectNodes("ctr")
                users.ClearATC
                For Each pNode In pNodes
                    Set ctr = New Controller
                    ctr.ID = getAttr(pNode, "code", "?")
                    ctr.NetworkID = getAttr(pNode, "networkID", "000000")
                    ctr.Frequency = getAttr(pNode, "freq", "199.98")
                    ctr.name = getAttr(pNode, "name", "???")
                    ctr.Rating = getAttr(pNode, "rating", "Observer")
                    ctr.FacilityType = getAttr(pNode, "type", "Center")
                    
                    'Add to the list
                    users.AddController ctr
                Next
                
                'Update the ATC list
                users.UpdateATCList
                If Not frmMain.SSTab1.TabEnabled(2) Then
                    frmMain.SSTab1.TabCaption(2) = info.Network + " Air Traffic Control"
                    frmMain.SSTab1.TabEnabled(2) = True
                    frmMain.SSTab1.TabVisible(2) = True
                End If
            
            Case "runways"
                Set pNode = cmdNode.selectSingleNode("runways")
                If (pNode Is Nothing) Then Exit Sub
                Set pNode = pNode.selectSingleNode("runway")
                If (pNode Is Nothing) Then Exit Sub
            
                'Display runway info
                freq = getChild(pNode, "freq", "")
                Msg = "Runway " + getAttr(pNode, "name", "?") + " at " + getAttr(pNode, "icao", "?") + _
                    ": " + Format(CInt(getAttr(pNode, "length", "0")), "##,##0") + " feet, " + _
                    Format(CInt(getAttr(pNode, "hdg", "0")), "000") + " degrees"
                If (freq <> "") Then Msg = Msg + " ILS: " + freq
                ShowMessage Msg, SYSMSGCOLOR
            
                'If we have a frequency, update the NAV1 radio
                If (freq <> "") Then
                    SetNAV1 freq, CInt(getAttr(pNode, "hdg", ""))
                    ShowMessage "NAV1 Radio set to " + freq, ACARSTEXTCOLOR
                End If
            
            Case "navaid"
                Dim navType As String
                Dim RadioCode As String
                
                'Get the navaid info
                Set pNode = cmdNode.selectSingleNode("navaid")
                If (pNode Is Nothing) Then Exit Sub
                Set pNode = pNode.selectSingleNode("navaid")
                If (pNode Is Nothing) Then Exit Sub
            
                'Set stuff based on navaid type. Override if necessary
                navType = UCase(getChild(pNode, "type", "VOR"))
                RadioCode = UCase(getChild(pNode, "radio", ""))
                If (navType = "NDB") Then RadioCode = "ADF"
                
                freq = getChild(pNode, "freq", "")
                Select Case RadioCode
                    Case "NAV1"
                        SetNAV1 freq, CInt(getChild(pNode, "hdg", "0"))
                    
                    Case "NAV2"
                        SetNAV2 freq
                        
                    Case "ADF"
                        SetADF1 freq
                End Select
            
            Case "airports"
                Dim a As Airport
            
                'Load the Airport Names/Codes
                Set pNode = cmdNode.selectSingleNode("airports")
                If Not (pNode Is Nothing) Then
                    Set pNodes = pNode.selectNodes("airport")
                    config.ClearAirports
                    For Each pNode In pNodes
                        Set a = New Airport
                        a.ICAO = getAttr(pNode, "icao")
                        a.IATA = getAttr(pNode, "iata")
                        a.name = Replace(getAttr(pNode, "name", a.ICAO), ",", "")
                        a.Latitude = CDbl(Replace(getAttr(pNode, "lat", "0"), ".", config.DecimalSeparator))
                        a.Longitude = CDbl(Replace(getAttr(pNode, "lng", "0"), ".", config.DecimalSeparator))
                        config.AddAirport a
                    Next
                
                    SetComboChoices frmMain.cboAirportD, config.AirportNames, "", "-"
                    SetComboChoices frmMain.cboAirportA, config.AirportNames, "", "-"
                    If config.ShowDebug Then ShowMessage "Updated Airport List, size=" + CStr(UBound(config.AirportNames) + 1), DEBUGTEXTCOLOR
                    config.SaveAirports
                End If
                
            Case "airlines"
                Set pNode = cmdNode.selectSingleNode("airlines")
                If Not (pNode Is Nothing) Then
                    Set pNodes = pNode.selectNodes("airline")
                    config.ClearAirlines
                    For Each pNode In pNodes
                        config.AddAirline getAttr(pNode, "code"), getAttr(pNode, "name")
                    Next
                    
                    SetComboChoices frmMain.cboAirline, config.AirlineNames, info.Airline.name, "-"
                    If config.ShowDebug Then ShowMessage "Updated Airline List", DEBUGTEXTCOLOR
                    config.SaveAirlines
                End If

            Case "pireps"
                Dim fr As FlightReport
            
                Set pNode = cmdNode.selectSingleNode("pireps")
                If (pNode Is Nothing) Then Exit Sub
                
                Load frmDraftPIREP
                Set pNodes = pNode.selectNodes("pirep")
                For Each pNode In pNodes
                    Set fr = New FlightReport
                    Set fr.Airline = config.GetAirline(getAttr(pNode, "airline", ""))
                    fr.FlightNumber = CInt(getAttr(pNode, "number", "001"))
                    fr.Leg = CInt(getAttr(pNode, "leg", "1"))
                    fr.EquipmentType = getChild(pNode, "eqType", "")
                    Set fr.airportD = config.GetAirport(getChild(pNode, "airportD", ""))
                    Set fr.AirportA = config.GetAirport(getChild(pNode, "airportA", ""))
                    fr.Remarks = getChild(pNode, "remarks", "")
                    
                    'Add to the list
                    frmDraftPIREP.AddFlight fr
                Next
                
                If (frmDraftPIREP.Size > 0) Then
                    frmDraftPIREP.Update
                    frmDraftPIREP.Show
                End If

        Case "info"
            'Load the equipment types
            Set pNode = cmdNode.selectSingleNode("info")
            If Not (pNode Is Nothing) Then
                'Get the private voice URL
                ID = getChild(pNode, "url", "")
                If (ID <> "") Then
                    config.PrivateVoiceURL = ID
                    If config.ShowDebug Then ShowMessage "Private Voice URL = " + ID, DEBUGTEXTCOLOR
                End If
            
                'Get the equipment types
                Set pNodes = pNode.selectNodes("eqtype")
                If (Not (pNodes Is Nothing) And (pNodes.Length > 0)) Then
                    config.ClearEquipment
                    For Each pNode In pNodes
                        config.AddEquipment Trim(pNode.Text)
                    Next
                
                    If config.ShowDebug Then ShowMessage "Updating Equipment List", DEBUGTEXTCOLOR
                    SetComboChoices frmMain.cboEquipment, config.EquipmentTypes, info.EquipmentType, "-"
                    config.SaveEquipment
                End If
            End If
            
        Case "charts"
            'Load the charts
            Set pNode = cmdNode.selectSingleNode("charts")
            If Not (pNode Is Nothing) Then
                With frmCharts
                    Set .ApproachCharts = New Charts
                    .ApproachCharts.AirportName = getAttr(pNode, "name", "")
                    .ApproachCharts.AirportCode = getAttr(pNode, "icao", "")
            
                    'Load the chart names
                    Set pNodes = pNode.selectNodes("chart")
                    For Each pNode In pNodes
                        .ApproachCharts.addChart getAttr(pNode, "name", "CHART"), CInt(getAttr(pNode, "id", "0"))
                    Next
                
                    .brwChart.Navigate "about:blank"
                    .cboAirport.Clear
                    .cboAirport.AddItem .ApproachCharts.AirportComboEntry
                    .cboAirport.ListIndex = 0
                    .ApproachCharts.Update
                    .Show
                    
                    If config.ShowDebug Then ShowMessage "Received charts for " + .ApproachCharts.AirportCode, DEBUGTEXTCOLOR
                End With
            End If
        End Select
    Next
End Sub

Attribute VB_Name = "RequestReceiver"
Option Explicit

Public Const XMLRESPONSEROOT = "ACARSResponse"

Private Function getChild(node As IXMLDOMNode, name As String, Optional defVal As String)
    Dim e As IXMLDOMNode

    Set e = node.selectSingleNode(name)
    If (e Is Nothing) Then
        getChild = defVal
    Else
        getChild = Trim(e.Text)
    End If
End Function

Private Function getAttr(node As IXMLDOMNode, name As String, Optional defVal As String)
    Dim e As IXMLDOMAttribute
    
    Set e = node.Attributes.getNamedItem(name)
    If (e Is Nothing) Then
        getAttr = defVal
    Else
        getAttr = Trim(e.Text)
    End If
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
        ShowMessage msgText, DEBUGTEXTCOLOR
        Exit Sub
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
        
            'Get the request ID
            ReqID = Val("&H" + getAttr(cmd, "id", "0"))
            
            'Take appropriate action based on request ID specified in the error.
            Select Case ReqID
                Case info.AuthReqID
                    frmMain.CloseACARSConnection False
                    MsgBox "The following error occurred while attempting to connect to the ACARS server:" & vbCrLf & vbCrLf & getChild(cmd, "error", "?"), vbOKOnly Or vbExclamation, "Error"
                Case info.PIREPReqID
                    MsgBox "The following error occurred while attempting to file the Flight Report:" & vbCrLf & vbCrLf & getChild(cmd, "error", "?"), vbOKOnly Or vbExclamation, "Error"
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
    ShowMessage "Error Processing ACARS Response - (Line " + CStr(Erl) + ") " + Error$(err), ACARSERRORCOLOR
    Resume ExitSub
End Sub

Private Sub ProcessACK(cmdNode As IXMLDOMNode)
    Dim ReqID As Long
    Static LastID As Long
    
    ReqID = Val("&H" + getAttr(cmdNode, "id", "0"))
    If (ReqID = LastID) Then Exit Sub
    
    If config.ShowDebug Then ShowMessage "Received ACK for " + Hex(ReqID), DEBUGTEXTCOLOR
    LastID = ReqID
            
    If ReqID = info.AuthReqID Then
        frmMain.sbMain.Panels(1).Text = "Status: Logged in to ACARS server"
        PlaySoundFile "notify_msg.wav"
        If Not config.DataUpdated Then
            RequestEquipment
            RequestAirports
            If config.SB3Support Then RequestPrivateVoiceURL
            config.DataUpdated = True
        End If
        
        RequestPilotList
        If (info.flightID <> 0) Then SendFlightInfo info
        ReqStack.Send
        DoEvents
    ElseIf ReqID = info.InfoReqID Then
        info.flightID = CLng(getChild(cmdNode, "flight_id", "0"))
        ShowMessage "Assigned Flight ID " + CStr(info.flightID), ACARSTEXTCOLOR

        'Save the flight ID to the registry for crash recovery.
        config.SaveFlightID info.flightID
    ElseIf ReqID = info.PIREPReqID Then
        info.PIREPFiled = True
        info.flightID = 0
        positions.Clear
        frmMain.cmdPIREP.Visible = False
        frmMain.cmdPIREP.Enabled = False
        info.FlightData = False
        MsgBox "Flight Report filed Successfully", vbInformation + vbOKOnly
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

    'Get the message info
    msgFrom = getChild(cmdNode, "from", "SYSTEM")
    msgTo = getChild(cmdNode, "to", "")
    msgText = getChild(cmdNode, "text", "")

    'Check if there is a "To" element. If so, it's a private message.
    If (msgTo = "") Then
        ShowMessage "<" & msgFrom & "> " & msgText, PUBLICCHATCOLOR
    Else
        ShowMessage "<" & msgFrom & ":PRIVATE> " & msgText, PRIVATECHATCOLOR
    End If

    'Play a sound if the option is on.
    If (config.PlaySound And (GetForegroundWindow <> frmMain.lngWindowHandle)) Then PlaySoundFile "notify_msg.wav"
End Sub

Private Sub ProcessDataResponse(cmdNode As IXMLDOMNode)
    Dim rspNodes As IXMLDOMNodeList
    Dim rspNode As IXMLDOMNode
    
    Dim pNodes As IXMLDOMNodeList
    Dim pNode As IXMLDOMNode
    Dim id As String
    Dim freq As String
    Dim Msg As String
    
    'Branch based on response type
    Set rspNodes = cmdNode.selectNodes("rsptype")
    For Each rspNode In rspNodes
        Select Case LCase(rspNode.Text)
    
            Case "pilotlist"
                Dim name As String
                
                If config.ShowDebug Then ShowMessage "Updating Pilot List", DEBUGTEXTCOLOR
                frmMain.lstPilots.Clear
                Set pNodes = cmdNode.selectSingleNode("pilotlist").selectNodes("Pilot")
                For Each pNode In pNodes
                    id = getAttr(pNode, "id", "")
                    name = getChild(pNode, "name", id)
                
                    'Check if the ID is ours
                    If (id <> "") Then
                        frmMain.lstPilots.AddItem id
                        If (UCase(id) = UCase(frmMain.txtPilotID.Text)) Then frmMain.txtPilotName.Text = name
                    End If
                Next

            Case "addpilots"
                Set pNodes = cmdNode.selectSingleNode("addpilots").selectNodes("Pilot")
                For Each pNode In pNodes
                    id = getAttr(pNode, "id", "")
                    If (id <> "") Then
                        frmMain.lstPilots.AddItem id
                        ShowMessage getChild(pNode, "name", id) + " logged into the ACARS server.", ACARSTEXTCOLOR
                    End If
                Next
            
            Case "delpilots"
                Dim x As Integer
                
                Set pNodes = cmdNode.selectSingleNode("delpilots").selectNodes("Pilot")
                For Each pNode In pNodes
                    id = getAttr(pNode, "id", "")
                    For x = 0 To frmMain.lstPilots.ListCount - 1
                        If frmMain.lstPilots.List(x) = id Then
                            frmMain.lstPilots.RemoveItem (x)
                            ShowMessage getChild(pNode, "name", id) + " logged out from the ACARS server.", ACARSTEXTCOLOR
                        End If
                    Next
                Next

        Case "pilot"
            Dim p As IXMLDOMNode
            
            'Get the pilot
            Set p = cmdNode.selectSingleNode("Pilot")
            If (p Is Nothing) Then Exit Sub
            
            'Build the message
            Msg = "Information for pilot " & getChild(p, "id", "") & ":" & vbCrLf
            Msg = Msg & "  Name: " & getChild(p, "name", "") & vbCrLf
            Msg = Msg & "  Time Online: " & getChild(p, "online_time", "") & vbCrLf
            Msg = Msg & "  Flight #: " & getChild(p, "flight_num", "") & vbCrLf
            Msg = Msg & "  Equipment: " & getChild(p, "equipment", "") & vbCrLf
            Msg = Msg & "  Cruise  Alt: " & getChild(p, "cruise_alt", "") & vbCrLf
            Msg = Msg & "  Departed: " & getChild(p, "dep_apt", "") & vbCrLf
            Msg = Msg & "  Arriving at: " & getChild(p, "arr_apt", "") & vbCrLf
            Msg = Msg & "  Route: " & getChild(p, "route", "") & vbCrLf
            Msg = Msg & "  Remarks: " & getChild(p, "remarks", "") & vbCrLf
            Msg = Msg & "  Alt MSL: " & getChild(p, "alt_msl", "") & vbCrLf
            Msg = Msg & "  Heading: " & getChild(p, "heading", "") & vbCrLf
            Msg = Msg & "  Air Speed: " & getChild(p, "air_speed", "") & vbCrLf
            Msg = Msg & "  Ground Speed: " & getChild(p, "ground_speed", "") & vbCrLf
            Msg = Msg & "  Flight Phase: " & getChild(p, "phase", "") & vbCrLf
            ShowMessage Msg, SYSMSGCOLOR
            
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
                setNAV1 freq, CInt(getAttr(pNode, "hdg", ""))
                ShowMessage "NAV1 Radio set to " + freq, ACARSTEXTCOLOR
            End If
            
        Case "navaid"
            'Get the navaid info
            Set pNode = cmdNode.selectSingleNode("navaid")
            If (pNode Is Nothing) Then Exit Sub
            Set pNode = pNode.selectSingleNode("navaid")
            If (pNode Is Nothing) Then Exit Sub
            
            'Set stuff based on navaid type
            freq = getChild(pNode, "freq", "")
            Select Case getChild(pNode, "radio", "")
                Case "nav1"
                    setNAV1 freq, CInt(getChild(pNode, "hdg", ""))
                    ShowMessage "NAV1 Radio set to " + freq, ACARSTEXTCOLOR
                    
                Case "nav2"
                    SetNAV2 freq
                    ShowMessage "NAV2 Radio set to " + freq, ACARSTEXTCOLOR
            End Select
            
        Case "airports"
            'Load the Airport Names/Codes
            Set pNode = cmdNode.selectSingleNode("airports")
            If Not (pNode Is Nothing) Then
                Set pNodes = pNode.selectNodes("airport")
                config.ClearAirports
                For Each pNode In pNodes
                    id = UCase(getAttr(pNode, "icao"))
                    config.AddAirport getAttr(pNode, "name") + " (" + id + ")", id
                Next
                
                SetComboChoices frmMain.cboAirportD, config.AirportNames
                SetComboChoices frmMain.cboAirportA, config.AirportNames
                If config.ShowDebug Then ShowMessage "Updated Airport List, size=" + CStr(UBound(config.AirportNames) + 1), DEBUGTEXTCOLOR
                config.SaveAirports
            End If

        Case "info"
            'Load the equipment types
            Set pNode = cmdNode.selectSingleNode("info")
            If Not (pNode Is Nothing) Then
                'Get the private voice URL
                id = getChild(pNode, "url", "")
                If (id <> "") Then
                    config.PrivateVoiceURL = id
                    If config.ShowDebug Then ShowMessage "Private Voice URL = " + id, DEBUGTEXTCOLOR
                End If
            
                'Get the equipment types
                Set pNodes = pNode.selectNodes("eqtype")
                If (Not (pNodes Is Nothing) And (pNodes.Length > 0)) Then
                    config.ClearEquipment
                    For Each pNode In pNodes
                        config.AddEquipment Trim(pNode.Text)
                    Next
                
                    If config.ShowDebug Then ShowMessage "Updating Equipment List", DEBUGTEXTCOLOR
                    SetComboChoices frmMain.cboEquipment, config.EquipmentTypes
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

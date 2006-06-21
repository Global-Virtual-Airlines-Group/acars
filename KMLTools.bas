Attribute VB_Name = "KMLTools"
Option Explicit

Private WarnMessages As Variant

Public Function createIcon(ByVal pal As Integer, ByVal x As Integer, ByVal y As Integer) As IXMLDOMElement
    Dim doc As New DOMDocument
    Dim re As IXMLDOMElement
    Dim e As IXMLDOMElement
    
    'Create the element
    Set re = doc.createNode(NODE_ELEMENT, "IconStyle", "")
    Set e = doc.createNode(NODE_ELEMENT, "Icon", "")
    
    'Set the fields
    AddXMLField e, "href", "root://icons/palette-" & CStr(pal) & ".png", False
    AddXMLField e, "x", CStr(x * 32), False
    AddXMLField e, "y", CStr(y * 32), False
    AddXMLField e, "h", "32", False
    AddXMLField e, "w", "32", False
    re.appendChild e
    
    'Return the element
    Set createIcon = re
End Function

Public Function createLookAt(ByVal lat As Double, ByVal lng As Double, ByVal alt As Integer, _
    ByVal hdg As Integer, ByVal tilt As Integer) As IXMLDOMElement
    Dim doc As New DOMDocument
    Dim e As IXMLDOMElement
    
    'Create the element
    Set e = doc.createNode(NODE_ELEMENT, "LookAt", "")
    
    'Set the fields
    AddXMLField e, "longitude", FormatNumber(lng, "##0.00000"), False
    AddXMLField e, "latitude", FormatNumber(lat, "##0.00000"), False
    AddXMLField e, "range", FormatNumber((0.3048 * alt), "##0.000"), False
    AddXMLField e, "heading", CStr(hdg), False
    AddXMLField e, "tilt", CStr(tilt), False
    
    'Return the element
    Set createLookAt = e
End Function

Public Function createAirport(a As Airport, ByVal desc As String) As IXMLDOMElement
    Dim doc As New DOMDocument
    Dim pe As IXMLDOMElement
    Dim se As IXMLDOMElement
    Dim ppe As IXMLDOMElement
    
    'Create the placemark element
    Set pe = doc.createNode(NODE_ELEMENT, "Placemark", "")
    AddXMLField pe, "description", desc, False
    AddXMLField pe, "name", a.name & " (" & a.ICAO & ")", False
    AddXMLField pe, "visibility", "1", False
    
    'Create the style element
    Set se = doc.createNode(NODE_ELEMENT, "Style", "")
    se.appendChild createIcon(2, 0, 1)
    AddXMLField se, "scale", "0.50", False
    pe.appendChild se
    
    'Create the Point element
    Set ppe = doc.createNode(NODE_ELEMENT, "Point", "")
    AddXMLField ppe, "coordinates", FormatNumber(a.Longitude, "##0.00000") & "," & _
        FormatNumber(a.Latitude, "###0.00000") & ",0", False
    pe.appendChild ppe
    pe.appendChild createLookAt(a.Latitude, a.Longitude, 4500, 1, 10)
    
    'Return the element
    Set createAirport = pe
End Function

Public Function createProgress(posData As Variant, RouteColor As KMLColor) As IXMLDOMElement
    Dim x As Integer
    Dim cPos As Variant
    Dim coords As String
    
    Dim doc As New DOMDocument
    Dim pe As IXMLDOMElement
    Dim se As IXMLDOMElement
    Dim lse As IXMLDOMElement
    
    'Create the placemark element
    Set pe = doc.createNode(NODE_ELEMENT, "Placemark", "")
    AddXMLField pe, "name", "Flight Route", False
    AddXMLField pe, "description", CStr(UBound(posData) + 1) & " Position Records", False
    AddXMLField pe, "visibility", "1", False
    
    'Create the style element
    Set se = doc.createNode(NODE_ELEMENT, "Style", "")
    AddXMLChild se, "LineStyle", "color", RouteColor.Text, False
    AddXMLField se, "width", "2.5", False
    AddXMLChild se, "PolyStyle", "color", RouteColor.Adjust(2.5).Text, False
    pe.appendChild se
    
    'Create the Line element
    Set lse = doc.createNode(NODE_ELEMENT, "LineString", "")
    AddXMLField lse, "extrude", "1", False
    AddXMLField lse, "altitudeMode", "relativeToGround", False
    AddXMLField lse, "tessellate", "1", False
    
    'Build the coordinates
    coords = ""
    For x = 0 To UBound(posData)
        Set cPos = posData(x)
        coords = coords & FormatNumber(cPos.Longitude, "##0.00000") & "," & _
            FormatNumber(cPos.Latitude, "###0.00000") & "," & CStr(cPos.AltitudeAGL) & _
            " " & vbCrLf
    Next
    
    'Save the coordinates
    AddXMLField lse, "coordinates", coords, False
    pe.appendChild lse

    'Return the element
    Set createProgress = pe
End Function

Public Function createPositionData(posData As Variant, Optional IsVisible As Boolean = False)
    Dim cPos As PositionData
    Dim x As Integer
    Dim geoPos As New GeoPosition
    
    Dim doc As New DOMDocument
    Dim fe As IXMLDOMElement
    Dim pe As IXMLDOMElement
    Dim se As IXMLDOMElement
    Dim ise As IXMLDOMElement
    Dim ppe As IXMLDOMElement
    
    'Init the warning messages
    If Not IsArray(WarnMessages) Then WarnMessages = Array("OK", "Over 250 Knots under 10,000 feet", _
        "High descent rate on Final Approach", "Abnormal Bank", "Abnormal Pitch", _
        "Abnormal G Forces")
    
    'Create the Folder element
    Set fe = doc.createNode(NODE_ELEMENT, "Folder", "")
    AddXMLField fe, "name", "Position Records", False
    setVisible fe, IsVisible
    
    'Create the position records
    For x = 0 To UBound(posData)
        Set cPos = posData(x)
        geoPos.setValue cPos.Latitude, cPos.Longitude
        Set pe = doc.createNode(NODE_ELEMENT, "Placemark", "")
        AddXMLField pe, "name", "Route Point #" & CStr(x + 1), False
        setVisible pe, IsVisible
        pe.appendChild createLookAt(cPos.Latitude, cPos.Longitude, (cPos.AltitudeMSL * 2) + 2000, _
            (cPos.Heading - 140), 15)
        AddXMLField pe, "description", createDataHTML(cPos), True
        AddXMLField pe, "snippet", geoPos.Text, False
        
        'Create style element
        Set se = doc.createNode(NODE_ELEMENT, "Style", "")
        
        'Set the icon if we have a warning or not
        If (cPos.WarningCode = 0) Then
            Set ise = createIcon(2, 0, 0)
        Else
            Set ise = createIcon(3, 2, 3)
        End If
        
        AddXMLField ise, "scale", "0.65", False
        AddXMLField ise, "heading", FormatNumber(cPos.Heading, "##0.00"), False
        se.appendChild ise
        pe.appendChild se
        
        'Create Point element
        Set ppe = doc.createNode(NODE_ELEMENT, "Point", "")
        If cPos.onGround Then
            AddXMLField ppe, "coordinates", Format(cPos.Longitude, "##0.00000") + "," + _
                FormatNumber(cPos.Latitude, "#0.00000"), False
            AddXMLField ppe, "altitudeMode", "clampedToGround", False
        ElseIf (cPos.AltitudeAGL < 1000) Then
            AddXMLField ppe, "coordinates", Format(cPos.Longitude, "##0.00000") + "," + _
                FormatNumber(cPos.Latitude, "#0.00000") + "," + FormatNumber(cPos.AltitudeAGL, "####0"), False
            AddXMLField ppe, "altitudeMode", "relativeToGround", False
        Else
            AddXMLField ppe, "coordinates", Format(cPos.Longitude, "##0.00000") + "," + _
                FormatNumber(cPos.Latitude, "#0.00000") + "," + FormatNumber(cPos.AltitudeMSL, "####0"), False
            AddXMLField ppe, "altitudeMode", "absolute", False
        End If
        
        'Add to folder
        pe.appendChild ppe
        fe.appendChild pe
    Next
    
    Set createPositionData = fe
End Function

Private Function createDataHTML(pos As PositionData) As String
    Dim results As String
    Dim geoPos As New GeoPosition
    Dim apValues As Variant
    Dim apCodes As Variant
    Dim x As Integer

    geoPos.Altitude = pos.AltitudeMSL
    geoPos.setValue pos.Latitude, pos.Longitude
    results = "<span class=""mapInfoBox"">Position: " + geoPos.Text + "<br /><br />" + _
        "Altitude: " + FormatNumber(pos.AltitudeMSL, "#,##0") + " feet"
    If ((pos.AltitudeAGL > 0) And (pos.AltitudeAGL < 2500)) Then results = results + _
        " (" + FormatNumber(pos.AltitudeAGL, "#,##0") + " feet AGL)"
    results = results + "<br />"
    If ((pos.Pitch < -1) Or (pos.Pitch > 5)) Then
        results = results + "Pitch: " + FormatNumber(pos.Pitch, "#0.0") + "<sup>o</sup>"
        If (Abs(pos.Bank) > 3) Then
            results = results + " "
        Else
            results = results + "<br />"
        End If
    End If
    
    If (Abs(pos.Bank) > 3) Then results = results + "Bank: " + FormatNumber(pos.Bank, "#0.0") + _
        "<sup>o</sup><br />"
    If (Abs(1 - pos.GForce) >= 0.1) Then results = results + "Acceleration: " + _
        FormatNumber(pos.GForce, "#0.000") + "G<br />"
        
    results = results + "Speed: " + FormatNumber(pos.AirSpeed, "##0") + " kts (GS: " + _
        FormatNumber(pos.GroundSpeed, "#,##0") + " kts)<br/>Heading: " + FormatNumber(pos.Heading, "000") + _
        " degrees<br />Vertical Speed:" + FormatNumber(pos.VerticalSpeed, "###0") + _
        " feet/min<br />N<sub>1</sub>: " + FormatNumber(pos.AverageN1, "##0.0") + _
        "%, N<sub>2</sub>: " + FormatNumber(pos.AverageN2, "##0.0") + "%<br/>Fuel Flow: " + _
        FormatNumber(pos.FuelFlow, "#,##0") + " lbs/hour<br />"
        
    If (pos.Flaps > 0) Then results = results + "Flaps: " + CStr(pos.Flaps) + "<sup>o</sup><br />"
    If pos.AfterBurner Then results = results + "<b><i>AFTERBURNER</i></b><br />"
    If pos.PUSHBACK Then results = results + "<b><i>PUSHBACK</i></b><br />"
    
    'Build Autopilot flags
    apValues = Array(pos.AP_HDG, pos.AP_NAV, pos.AP_GPS, pos.AP_APR, pos.AP_ALT)
    apCodes = Array("HDG", "NAV", "GPS", "APR", "ALT")
    If pos.AP_HDG Or pos.AP_NAV Or pos.AP_GPS Or pos.AP_APR Or pos.AP_ALT Then
        results = results + "Autopilot: "
        For x = 0 To UBound(apValues)
            If apValues(x) Then results = results + apCodes(x)
        Next
        
        results = results + "<br />"
    End If
    
    If pos.AT_IAS Then results = results + "Autothrottle: IAS<br />"
    If pos.AT_MCH Then results = results + "Autothrottle: MACH<br />"
    
    'Add warning message
    If (pos.WarningCode > 0) Then results = results + "<br /><span style=""color: #E02010"">" + _
        WarnMessages(pos.WarningCode) + "</span>"
    
    createDataHTML = results
End Function

Private Sub setVisible(entry As IXMLDOMElement, Optional IsVisible As Boolean = True)
    If IsVisible Then
        AddXMLField entry, "visibility", "1", False
    Else
        AddXMLField entry, "visibility", "0", False
    End If
End Sub

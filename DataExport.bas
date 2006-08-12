Attribute VB_Name = "DataExport"
Option Explicit

Public Sub WriteCSV(ByVal fName As String, posData As Variant)
    Dim fNum As Integer
    Dim x As Integer
    Dim cPos As Variant
    
    On Error GoTo FatalError
    fNum = FreeFile()
    Open fName + ".csv" For Output As #fNum
    Print #fNum, "; Date / Time,Latitude,Longitude,Altitude,Heading,Air Speed," & _
            "Ground Speed,Vertical Speed,N1,N2,Bank,Pitch,Flaps,Wind Speed,WindHdg," & _
            "Fuel Flow,Gs,AOA,Frame Rate"
            
    'Write position data
    For x = 0 To UBound(posData)
        Set cPos = posData(x)
        Print #fNum, FormatDateTime(cPos.DateTime.UTCTime, "mm/dd/yyyy hh:nn:ss") + "," + _
            FormatNumber(cPos.Latitude, "#0.00000") + "," + FormatNumber(cPos.Longitude, "##0.00000") _
            + "," + CStr(cPos.AltitudeMSL) + "," + CStr(cPos.Heading), "," + CStr(cPos.Airspeed) _
            + "," + CStr(cPos.GroundSpeed) + "," + CStr(cPos.VerticalSpeed) + "," + _
            FormatNumber(cPos.AverageN1, "##0.0") + "," + FormatNumber(cPos.AverageN2, "##0.0") _
            + "," + FormatNumber(cPos.Bank, "#0.000") + "," + FormatNumber(cPos.Pitch, "#0.000") _
            + "," + CStr(cPos.Flaps) + "," + CStr(cPos.WindSpeed) + "," + CStr(cPos.WindHeading) _
            + "," + CStr(cPos.FuelFlow) + "," + FormatNumber(cPos.GForce, "#0.000") + "," + _
            FormatNumber(cPos.AngleOfAttack, "#0.000") + "," + CStr(cPos.FrameRate)
    Next
    
    Close #fNum
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot save CSV Flight Data!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub
    
End Sub

Public Sub WriteKML(ByVal fName As String, posData As Variant, info As FlightData)
    Dim fNum As Integer
    Dim lColor As New KMLColor
    Dim apInfo As String
    
    Dim xdoc As New DOMDocument
    Dim pi As IXMLDOMProcessingInstruction
    Dim root As IXMLDOMElement
    Dim doc As IXMLDOMElement
    Dim shaData As String
    
    'Set the line color
    lColor.SetColors 68, 184, 244, 160
    
    'Build the document root
    Set root = xdoc.createNode(NODE_ELEMENT, "kml", "")
    Set doc = xdoc.createNode(NODE_ELEMENT, "Document", "")
    AddXMLField doc, "name", info.Remarks, False
    root.appendChild doc
    xdoc.appendChild root
    
    'Build the encoder
    Set pi = xdoc.createProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")
    xdoc.appendChild pi
    
    'Build the departure/takeoff data
    apInfo = "Departed from " & info.airportD.Name & " at " & CStr(info.TakeoffTime.LocalTime) & "<br />" & _
        CStr(info.TakeoffSpeed) & " knots, " & CStr(info.TakeoffN1) & "% N<sub>1</sub>, " & _
        CStr(info.TakeoffWeight) & " lbs total, " & CStr(info.TakeoffFuel) & " lbs fuel<br />"
    doc.appendChild createAirport(info.airportD, apInfo)
    
    'Add the flight data
    doc.appendChild createProgress(posData, lColor)
    doc.appendChild createPositionData(posData, False)
    
    'Build the arrival/landing data
    apInfo = "Landed at " & info.AirportA.Name & " at " & CStr(info.LandingTime.LocalTime) & "<br />" & _
        CStr(info.LandingSpeed) & " knots, " & CStr(info.LandingVSpeed) & " feet/min, " & _
        CStr(info.LandingN1) & "% N<sub>1</sub>, " & CStr(info.LandingWeight) & " lbs total, " & _
        CStr(info.LandingFuel) & " lbs fuel<br />"
    doc.appendChild createAirport(info.AirportA, apInfo)
    
    'Save the XML data
    Dim sha As New SHA256
    shaData = sha.SHA256(doc.XML, "***REMOVED***")
    
    'Write the document to disk
    fNum = FreeFile()
    Open fName + ".kml" For Output As #fNum
    Print #fNum, doc.XML
    Close #fNum
    
    'Write the SHA256 hash to disk
    fNum = FreeFile()
    Open fName + ".sha" For Output As #fNum
    Print #fNum, shaData
    Close #fNum
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot save Google Earth Flight Data", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub
    
End Sub

Public Sub BuildPackage(ByVal Name As String, Optional deleteFiles As Boolean = False)
    Dim z As New ZipClass

    z.AddFile Name & ".csv"
    z.AddFile Name & ".kml"
    z.AddFile Name & ".sha"
    z.Comment = "ACARS Flight " & Name
    z.WriteZip Name & ".zip"
    
    'Delete the files if requested
    If deleteFiles Then
        On Error Resume Next
        Kill Name & ".csv"
        Kill Name & ".kml"
        Kill Name & ".sha"
    End If
End Sub

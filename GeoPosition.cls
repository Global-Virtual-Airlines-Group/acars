VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeoPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public degreesLAT As Integer
Public degreesLNG As Integer
Public MinutesLAT As Integer
Public MinutesLNG As Integer
Public SecondsLAT As Integer
Public SecondsLNG As Integer
Public Altitude As Integer

Public Sub setValue(ByVal lat As Double, ByVal lng As Double)
    Me.degreesLAT = Fix(lat)
    Me.degreesLNG = Fix(lng)
    lat = lat - Me.degreesLAT
    lng = lng - Me.degreesLNG
    Me.MinutesLAT = Fix(lat * 60#)
    Me.MinutesLNG = Fix(lng * 60#)
    lat = lat - (Me.MinutesLAT / 60)
    lng = lng - (Me.MinutesLNG / 60)
    Me.SecondsLAT = (lat * 3600#)
    Me.SecondsLNG = (lng * 3600#)
End Sub

Property Get Latitude() As Double
    Latitude = degreesLAT + (MinutesLAT / 60#) + (SecondsLAT / 3600#)
End Property

Property Get Longitude() As Double
    Longitude = degreesLNG + (MinutesLNG / 60#) + (SecondsLNG / 3600#)
End Property

Property Get Text() As String
    Dim results As String
    
    'Build latitude
    results = CStr(Abs(degreesLAT)) + "<sup>o</sup> " + CStr(MinutesLAT) + _
        "&#39; " + CStr(SecondsLAT) + " &quot; "
    If (degreesLAT < 0) Then
        results = results + "S "
    Else
        results = results + "N "
    End If
    
    'Build longitude
    results = results + CStr(Abs(degreesLNG)) + "<sup>o</sup> " + CStr(MinutesLNG) + _
        "&#39; " + CStr(SecondsLNG) + " &quot; "
    If (degreesLNG < 0) Then
        results = results + "W"
    Else
        results = results + "E"
    End If
    
    Text = results
End Property

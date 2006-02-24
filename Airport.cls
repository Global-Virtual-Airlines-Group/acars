VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Airport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Conversion constants
Private Const RADIAN_DEGREE = 57.29578
Private Const DEGREE_MILES = 69.16

Private iataCode As String
Private icaoCode As String

Public Name As String

Public Latitude As Double
Public Longitude As Double

Property Let IATA(code As String)
    code = UCase(code)
    If (Len(code) > 3) Then code = Left(code, 3)
    iataCode = code
End Property

Property Let ICAO(code As String)
    code = UCase(code)
    If (Len(code) > 4) Then code = Left(code, 4)
    icaoCode = code
End Property

Property Get IATA() As String
    IATA = iataCode
End Property

Property Get ICAO() As String
    ICAO = icaoCode
End Property

Private Function ArcCos(ByVal x As Double) As Double
    ' Inverse Cosine
    If Abs(x) = 1 Then
        ArcCos = Atn(1) * (1 - Sgn(x)) * 4
    Else
        ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    End If
End Function

Function DistanceTo(lat As Double, lon As Double) As Integer
    Dim lat1 As Double, lat2 As Double
    Dim p1 As Double, p2 As Double
    Dim lngDiff As Double, distD As Double
    
    'Convert latitude to radians
    lat1 = Latitude / RADIAN_DEGREE
    lat2 = lat / RADIAN_DEGREE
    
    'Get the longitude difference in radians
    lngDiff = Abs(Longitude - lon) / RADIAN_DEGREE

    'Do the math - this makes my head hurt
    p1 = Sin(lat1) * Sin(lat2)
    p2 = Cos(lat1) * Cos(lat2) * Cos(lngDiff)
    distD = ArcCos(p1 + p2) * RADIAN_DEGREE

    'Convert to miles and return
    DistanceTo = Round(distD * DEGREE_MILES)
End Function

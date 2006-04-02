VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlightReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public flightCode As String
Public Leg As Byte
Public EquipmentType As String
Public Network As String
Public Remarks As String

Public airportD As Airport
Public AirportA As Airport

Property Get Description() As String
    If ((airportD Is Nothing) Or (AirportA Is Nothing)) Then Exit Property
    
    Description = airportD.Name + " (" + airportD.ICAO + ") - " + _
        AirportA.Name + " (" + AirportA.ICAO + ")"
End Property

Private Sub Class_Initialize()
    Network = "Offline"
End Sub

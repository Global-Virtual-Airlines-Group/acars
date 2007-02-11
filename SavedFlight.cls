VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SavedFlight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FlightInfo As FlightData
Public XML As String

Public AircraftAGL As Integer

Private posData As New OfflinePositionData

Property Get FlightID() As Integer
    If (FlightInfo Is Nothing) Then
        FlightID = 0
    Else
        FlightID = FlightInfo.FlightID
    End If
End Property

Public Sub AddPosition(cPos As PositionData)
    posData.AddPosition cPos
End Sub

Property Get Positions() As Variant
    Positions = posData.Queue
End Property

Property Get HasData() As Boolean
    HasData = posData.HasData
End Property

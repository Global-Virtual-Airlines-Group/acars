VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlightData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Flight information
Public flightID As Long

Public EquipmentType As String
Public FlightNumber As String
Public CruiseAltitude As String
Public AirportA As String
Public AirportD As String
Public Route As String
Public Remarks As String
Public Network As String

Public Phase As String
Public FSVersion As Integer
Public Offline As Boolean

'Log time at different sim rates
Private Time1X As Long
Private Time2X As Long
Private Time4X As Long

'Track request numbers for special messages
Private Auth_ReqID As Long
Private Info_ReqID As Long
Private PIREP_ReqID As Long

'Flight Start Data
Public StartTime As Date
Public TimeOut As Date
Public TaxiOutTime As Date
Public TaxiFuel As Long
Public TaxiWeight As Long

'Takeoff data
Public TimeOff As Date
Public TakeoffTime As Date
Public TakeoffSpeed As Integer
Public TakeoffFuel As Long
Public TakeoffWeight As Long
Public TakeoffN1 As Double

'Landing Data
Public LandingTime As Date
Public LandingSpeed As Integer
Public LandingVSpeed As Integer
Public LandingFuel As Long
Public LandingWeight As Long
Public LandingN1 As Double

'Flight end data
Public TaxiInTime As Date
Public GateTime As Date
Public GateFuel As Long
Public GateWeight As Long
Public ShutdownTime As Date
Public ShutdownFuel As Long
Public ShutdownWeight As Long

'State variables
Public InFlight As Boolean
Public FlightData As Boolean
Public PIREPFiled As Boolean

Sub Class_Initialize()
    Phase = "N/A"
End Sub

'Start a new flight
Sub StartFlight()
    InFlight = True
    FlightData = False
    PIREPFiled = False
    Phase = "Preflight"
    StartTime = Now
End Sub

'Updates flight time
Sub UpdateFlightTime(ByVal simRate As Integer, ByVal interval As Integer)
    If info.Phase <> "Airborne" Then Exit Sub

    interval = interval * simRate
    Select Case (simRate)
        Case 1
            Time1X = Time1X + interval
        Case 2
            Time2X = Time2X + interval
        Case 4
            Time4X = Time4X + interval
    End Select
End Sub

Property Get TimeAt1X() As Integer
    TimeAt1X = Time1X / 60000
End Property

Property Get TimeAt2X() As Integer
    TimeAt2X = Time2X / 60000
End Property

Property Get TimeAt4X() As Integer
    TimeAt4X = Time4X / 60000
End Property

Property Get AuthReqID() As Long
    AuthReqID = Auth_ReqID
End Property

Property Let AuthReqID(id As Long)
    Auth_ReqID = id
    If config.ShowDebug Then ShowMessage "AuthReq = " + Hex(id), DEBUGTEXTCOLOR
End Property

Property Get InfoReqID() As Long
    InfoReqID = Info_ReqID
End Property

Property Let InfoReqID(id As Long)
    Info_ReqID = id
    If config.ShowDebug Then ShowMessage "InfoReq = " + Hex(id), DEBUGTEXTCOLOR
End Property

Property Get PIREPReqID() As Long
    PIREPReqID = PIREP_ReqID
End Property

Property Let PIREPReqID(id As Long)
    PIREP_ReqID = id
    If config.ShowDebug Then ShowMessage "PIREPReq = " + Hex(id), DEBUGTEXTCOLOR
End Property

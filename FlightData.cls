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
Public FlightID As Long
Public CheckRide As Boolean

Public EquipmentType As String
Public FlightNumber As Integer
Public FlightLeg As Integer
Public CruiseAltitude As String
Public Airline As Airline
Public AirportA As Airport
Public airportD As Airport
Public AirportL As Airport
Public Route As String
Public Remarks As String
Public Network As String

Public FlightPhase As Integer
Private PhaseNames As Variant

Public FSVersion As Integer
Public Offline As Boolean

'Log time at different sim rates
Private Time0X As Long
Private Time1X As Long
Private Time2X As Long
Private Time4X As Long

'Track request numbers for special messages
Private Auth_ReqID As Long
Private Info_ReqID As Long
Private Sched_ReqID As Long

'Flight Start Data
Public StartTime As New UTCDate
Public TaxiOutTime As New UTCDate
Public TaxiFuel As Long
Public TaxiWeight As Long

'Takeoff data
Public TakeoffTime As New UTCDate
Public TakeoffSpeed As Integer
Public TakeoffFuel As Long
Public TakeoffWeight As Long
Public TakeoffN1 As Double

'Landing Data
Public LandingTime As New UTCDate
Public LandingSpeed As Integer
Public LandingVSpeed As Integer
Public LandingG As Double
Public LandingFuel As Long
Public LandingWeight As Long
Public LandingN1 As Double

'Flight end data
Public TaxiInTime As New UTCDate
Public GateTime As New UTCDate
Public GateFuel As Long
Public GateWeight As Long
Public ShutdownTime As New UTCDate
Public ShutdownFuel As Long
Public ShutdownWeight As Long

'State variables
Public InFlight As Boolean
Public FlightData As Boolean
Public ScheduleVerified As Boolean
Public PIREPFiled As Boolean

Sub Class_Initialize()
    PhaseNames = Array("N/A", "Pre-Flight", "Pushback", "Taxi Out", "Takeoff", "Airborne", _
        "Landed", "Taxi In", "At Gate", "Shutdown", "Complete", "Aborted", "Error", _
        "PIREP File")
    Network = "Offline"
    FlightPhase = UNKNOWN
    InFlight = False
    FlightData = False
    PIREPFiled = False
    ScheduleVerified = False
    Offline = False
    Set Airline = config.GetAirline(LoadResString(101))
    FlightLeg = 1
End Sub

'Start a new flight
Sub StartFlight(Optional isOnline As Boolean = True)
    InFlight = True
    FlightData = False
    PIREPFiled = False
    Offline = Not isOnline
    FlightPhase = PREFLIGHT
    StartTime.SetNow
    Time0X = 0
    Time1X = 0
    Time2X = 0
    Time4X = 0
End Sub

Property Get PhaseName() As String
    If ((FlightPhase < 0) Or (FlightPhase > UBound(PhaseNames))) Then FlightPhase = 0
    PhaseName = PhaseNames(FlightPhase)
End Property

Property Get flightCode() As String
    Dim fCode As String
    
    If (Airline Is Nothing) Then
        fCode = "DAL"
    Else
        fCode = Airline.code
    End If
    
    flightCode = fCode + CStr(FlightNumber)
End Property

'Updates flight time
Function UpdateFlightTime(ByVal simRate As Integer, ByVal interval As Long) As Date
    Dim tmpS As Integer

    If (simRate > 0) Then interval = interval * simRate
    Select Case (simRate)
        Case 0
            Time0X = Time0X + interval
        Case 1
            Time1X = Time1X + interval
        Case 2
            Time2X = Time2X + interval
        Case 4
            Time4X = Time4X + interval
    End Select
    
    'Return a time value
    tmpS = FlightTime Mod 3600
    UpdateFlightTime = TimeSerial(Fix(FlightTime / 3600), Fix(tmpS / 60), tmpS Mod 60)
End Function

Property Get TimePaused() As Long
    TimePaused = Time0X / 1000
End Property

Property Let TimePaused(ByVal ms As Long)
    Time0X = ms * 1000
End Property

Property Get TimeAt1X() As Long
    TimeAt1X = Time1X / 1000
End Property

Property Let TimeAt1X(ByVal ms As Long)
    Time1X = ms * 1000
End Property

Property Get TimeAt2X() As Long
    TimeAt2X = Time2X / 1000
End Property

Property Let TimeAt2X(ByVal ms As Long)
    Time2X = ms * 1000
End Property

Property Get TimeAt4X() As Long
    TimeAt4X = Time4X / 1000
End Property

Property Let TimeAt4X(ByVal ms As Long)
    Time4X = ms * 1000
End Property

Property Get FlightTime() As Long
    FlightTime = CLng(Time1X + Time2X + Time4X) / 1000
End Property

Property Get AuthReqID() As Long
    AuthReqID = Auth_ReqID
End Property

Property Let AuthReqID(ByVal ID As Long)
    Auth_ReqID = ID
    If config.ShowDebug Then ShowMessage "AuthReq = " + Hex(ID), DEBUGTEXTCOLOR
End Property

Property Get InfoReqID() As Long
    InfoReqID = Info_ReqID
End Property

Property Let InfoReqID(ByVal ID As Long)
    Info_ReqID = ID
    If config.ShowDebug Then ShowMessage "InfoReq = " + Hex(ID), DEBUGTEXTCOLOR
End Property

Property Get SchedReqID() As Long
    SchedReqID = Sched_ReqID
End Property

Property Let SchedReqID(ByVal ID As Long)
    Sched_ReqID = ID
    If config.ShowDebug Then ShowMessage "SchedReq = " + Hex(ID), DEBUGTEXTCOLOR
End Property

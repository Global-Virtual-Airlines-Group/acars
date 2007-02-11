VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Date/time of flight data
Public DateTime As New UTCDate
Public phase As String
Public FrameRate As Integer

'Latitude/longitude
Public Latitude As Double
Public Longitude As Double
    
'Current parameters
Public Mach As Double
Public Airspeed As Integer
Public GroundSpeed As Integer
Public VerticalSpeed As Integer
Public TouchdownSpeed As Integer
Public Pitch As Double
Public Bank As Double
Public Heading As Integer
Public AltitudeMSL As Long
Public AltitudeAGL As Long
Public Flaps As Integer
Public Spoilers As Boolean
Public GearDown As Boolean
Public EnginesStarted As Boolean
Public WindSpeed As Integer
Public WindHeading As Integer
Public GForce As Double
Public AngleOfAttack As Double

'Flight state flags
Public Paused As Boolean
Public Touchdown As Boolean
Public Slewing As Boolean
Public PUSHBACK As Boolean
Public onGround As Boolean
Public AfterBurner As Boolean
Public Parked As Boolean
Public simRate As Integer
Public CockpitView As Boolean
Public Stall As Boolean
Public Overspeed As Boolean
Public Crashed As Boolean

'Autopilot flags
Public AP_HDG As Boolean
Public AP_NAV As Boolean
Public AP_GPS As Boolean
Public AP_APR As Boolean
Public AP_ALT As Boolean
Public AT_IAS As Boolean
Public AT_MCH As Boolean

'Engine parameters
Private engN1(3) As Double
Private engN2(3) As Double
Private engThrottle(3) As Integer
Private engFlow(3) As Long
Private engCount As Integer

'Fuel/Weight state
Public Fuel As Long
Public weight As Long

'Gauge state
Public ACARSConnected As Boolean
Public ACARSPhase As Byte
Public SwitchCOM As Boolean

Public Sub setN1(ByVal engNum As Byte, n1 As Double)
    engN1(engNum) = n1
End Sub

Public Sub setN2(ByVal engNum As Byte, n2 As Double)
    engN2(engNum) = n2
End Sub

Public Sub setThrottle(ByVal engNum As Byte, tr As Integer)
    engThrottle(engNum) = tr
End Sub

Public Sub setFuelFlow(ByVal engNum As Byte, flow As Long)
    engFlow(engNum) = flow
End Sub

Public Function getN1(ByVal engNum As Byte) As Double
    getN1 = engN1(engNum)
End Function

Public Function getN2(ByVal engNum As Byte) As Double
    getN2 = engN2(engNum)
End Function

Public Function getThrottle(ByVal engNum As Byte) As Integer
    getThrottle = engThrottle(engNum)
End Function

Public Function getFuelFlow(ByVal engNum As Byte) As Long
    getFuelFlow = engFlow(engNum)
End Function

Property Let EngineCount(ByVal eCount As Integer)
    engCount = eCount
End Property

Property Get EngineCount() As Integer
    EngineCount = engCount
End Property

Property Get AverageN1() As Double
    Dim x As Integer
    Dim TotalN1 As Double

    For x = 0 To (engCount - 1)
        TotalN1 = TotalN1 + engN1(x)
    Next
    
    AverageN1 = TotalN1 / engCount
End Property

Property Get AverageN2() As Double
    Dim x As Integer
    Dim TotalN2 As Double

    For x = 0 To (engCount - 1)
        TotalN2 = TotalN2 + engN2(x)
    Next
    
    AverageN2 = TotalN2 / engCount
End Property

Property Get AverageThrottle() As Integer
    Dim x As Integer
    Dim totalT As Long

    For x = 0 To (engCount - 1)
        totalT = totalT + engThrottle(x)
    Next
    
    AverageThrottle = totalT / engCount
End Property

Property Get FuelFlow() As Long
    Dim x As Integer
    Dim totalFlow As Long
    
    For x = 0 To (engCount - 1)
        totalFlow = totalFlow + engFlow(x)
    Next
    
    FuelFlow = totalFlow
End Property

Private Sub Class_Initialize()
    AP_HDG = False
    AP_NAV = False
    AP_GPS = False
    AP_APR = False
    AP_ALT = False
    AT_IAS = False
    AT_MCH = False
    Paused = False
    Slewing = False
    Spoilers = False
    GearDown = False
    onGround = False
    Touchdown = False
    EnginesStarted = False
    Parked = False
    AfterBurner = False
    Stall = False
    Overspeed = False
    Crashed = False
End Sub

Property Get Flags() As Long
    Dim tmpFlags As Long
    
    tmpFlags = 0
    
    'Build the flags
    If Paused Then tmpFlags = tmpFlags Or FLIGHT_PAUSED
    If Touchdown Then tmpFlags = tmpFlags Or FLIGHT_TOUCHDOWN
    If Parked Then tmpFlags = tmpFlags Or FLIGHT_PARKED
    If onGround Then tmpFlags = tmpFlags Or FLIGHT_ONGROUND
    If Spoilers Then tmpFlags = tmpFlags Or FLIGHT_SP_ARM
    If GearDown Then tmpFlags = tmpFlags Or FLIGHT_GEAR_DOWN
    If AfterBurner Then tmpFlags = tmpFlags Or FLIGHT_AFTERBURNER
    If Overspeed Then tmpFlags = tmpFlags Or FLIGHT_OVERSPEED
    If Stall Then tmpFlags = tmpFlags Or FLIGHT_STALL
    If Crashed Then tmpFlags = tmpFlags Or FLIGHT_CRASH
    If AP_NAV Then tmpFlags = tmpFlags Or FLIGHT_AP_NAV
    If AP_GPS Then tmpFlags = tmpFlags Or FLIGHT_AP_GPS
    If AP_HDG Then tmpFlags = tmpFlags Or FLIGHT_AP_HDG
    If AP_APR Then tmpFlags = tmpFlags Or FLIGHT_AP_APR
    If AP_ALT Then tmpFlags = tmpFlags Or FLIGHT_AP_ALT
    If AT_IAS Then tmpFlags = tmpFlags Or FLIGHT_AT_IAS
    If AT_MCH Then tmpFlags = tmpFlags Or FLIGHT_AT_MACH
    Flags = tmpFlags
End Property


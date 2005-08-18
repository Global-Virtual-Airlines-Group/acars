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
Public DateTime As PositionDate
Public Phase As String

'Latitude/longitude
Public Latitude As Double
Public Longitude As Double
    
'Current parameters
Public Mach As Double
Public AirSpeed As Integer
Public GroundSpeed As Integer
Public VerticalSpeed As Integer
Public Heading As Integer
Public AltitudeMSL As Long
Public AltitudeAGL As Long
Public Flaps As Integer
Public Spoilers As Boolean
Public GearDown As Boolean
Public EnginesStarted As Boolean

'Flight state flags
Public Paused As Boolean
Public Slewing As Boolean
Public onGround As Boolean
Public Parked As Boolean
Public simRate As Integer

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
Private engCount As Integer

'Fuel/Weight state
Public Fuel As Long
Public Weight As Long

Public Sub Class_Initialize()
    Set DateTime = New PositionDate
End Sub

Public Sub setN1(ByVal engNum As Integer, n1 As Double)
    engN1(engNum) = n1
End Sub

Public Sub setN2(ByVal engNum As Integer, n2 As Double)
    engN2(engNum) = n2
End Sub

Public Sub setThrottle(ByVal engNum As Integer, tr As Integer)
    engThrottle(engNum) = tr
End Sub

Public Function getN1(ByVal engNum As Integer) As Double
    getN1 = engN1(engNum)
End Function

Public Function getN2(ByVal engNum As Integer) As Double
    getN2 = engN2(engNum)
End Function

Public Function getThrottle(ByVal engNum As Byte) As Integer
    getThrottle = engThrottle(engNum)
End Function

Property Let EngineCount(ByVal eCount As Integer)
    engCount = eCount
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

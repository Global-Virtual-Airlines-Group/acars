VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AircraftInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PMDG737 As Boolean
Public PMDG747 As Boolean
Public LDS767 As Boolean
Public Payne7X7 As Boolean

Public ZeroFuelWeight As Long
Public MaxGrossWeight As Long
Public HasAfterburner As Boolean
Private acAGL As Integer

Public FSPath As String
Private airP As String
Private airF As String

Public Code As String

'Aircraft parameters
Public FuelWeight As Double
Public EngineType As Integer
Public EngineCount As Integer

'FSUIPC Fuel Tank Offsets
Private TankOfs As Variant
Private TankCOfs As Variant

Private TankSizes(10) As Long

Public Sub Class_Initialize()
    PMDG737 = False
    PMDG747 = False
    LDS767 = False
    Payne7X7 = False
    HasAfterburner = False
    acAGL = -1
End Sub

Property Let AirPath(ByVal path As String)
    Dim x As Integer
    
    x = InStr(StrReverse(path), "\")
    airF = Mid(path, Len(path) - (x - 2))
    airP = Left(path, Len(path) - (x - 1))
End Property

Property Get AirPath() As String
    AirPath = airP + airF
End Property

Property Get AIRFile() As String
    AIRFile = airF
End Property

Property Get CFGFile()
    CFGFile = airP + "aircraft.cfg"
End Property

Property Get AircraftPath() As String
    AircraftPath = airP
End Property

Property Get BaseAGL() As Integer
    BaseAGL = acAGL
End Property

Property Let BaseAGL(ByVal agl As Integer)
    If (acAGL = -1) Then acAGL = agl
End Property

Property Get tankOffsets() As Variant
    tankOffsets = TankOfs
End Property

Property Get TankCOffsets() As Variant
    TankCOffsets = TankCOfs
End Property

Public Sub AddTankOffset(ByVal ofs As Long)
    Dim x As Integer
    
    If IsEmpty(TankOfs) Then
        ReDim TankOfs(0)
    Else
        For x = 0 To UBound(TankOfs)
            If (TankOfs(x) = ofs) Then Exit Sub
        Next
    
        ReDim Preserve TankOfs(UBound(TankOfs) + 1)
    End If
        
    TankOfs(UBound(TankOfs)) = ofs
End Sub

Public Sub AddTankCapacityOffset(ByVal ofs As Long)
    Dim x As Integer

    If IsEmpty(TankCOfs) Then
        ReDim TankCOfs(0)
    Else
        For x = 0 To UBound(TankCOfs)
            If (TankCOfs(x) = ofs) Then Exit Sub
        Next
    
        ReDim Preserve TankCOfs(UBound(TankCOfs) + 1)
    End If
        
    TankCOfs(UBound(TankCOfs)) = ofs
End Sub

Public Sub SetCapacity(ByVal tank As Integer, ByVal Size As Long)
    If ((tank >= 0) And (tank < 10)) Then TankSizes(tank) = CLng(Size * FuelWeight)
End Sub

Public Function TankCapacity(ByVal tank As Integer) As Long
    If ((tank >= 0) And (tank < 10)) Then TankCapacity = TankSizes(tank)
End Function

Public Function HasTank(ByVal tank As Integer) As Boolean
    HasTank = (TankCapacity(tank) > 0)
End Function

Property Get Tanks() As Variant
    Dim x As Integer
    Dim results As Variant
    
    For x = 0 To MAX_TANK
        If (TankCapacity(x) > 0) Then
            ReDim Preserve results(UBound(results) + 1)
            results(UBound(results)) = x
        End If
    Next
    
    Tanks = results
End Property

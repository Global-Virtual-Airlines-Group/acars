VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Failures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private engCount As Integer
Public MaxFailures As Integer
Public InstrumentFailureProbability As Integer
Public EquipmentFailureProbability As Integer

Public Force As Boolean

Public ADF As Boolean
Public Altimeter As Boolean
Public Attitude As Boolean
Public Airspeed As Boolean
Public COM1 As Boolean
Public COM2 As Boolean
Public Compass As Boolean
Public Fuel As Boolean
Public NAV1 As Boolean
Public NAV2 As Boolean
Public Transponder As Boolean
Public VerticalSpeed As Boolean

Private engFailure(1 To 4) As Boolean
Private revFailure(1 To 4) As Boolean
Public Flaps As Boolean
Public Spoilers As Boolean
Public Gear As Boolean
Public PitotHeat As Boolean

Public Sub Class_Initialize()
    Randomize
End Sub

Property Get IsInstrumentFailure() As Boolean
    Dim fprob As Single
    fprob = InstrumentFailureProbability * 0.016667 / 60
    IsInstrumentFailure = (Rnd < fprob)
End Property

Property Get IsEquipmentFailure() As Boolean
    Dim fprob As Single
    fprob = EquipmentFailureProbability * 0.016667 / 60
    IsEquipmentFailure = (Rnd < fprob)
End Property

Property Let EngineCount(ByVal eCount As Integer)
    engCount = eCount
End Property

Public Sub SetEngineFailure(ByVal eng As Byte, ByVal hasFailure As Boolean)
    If ((eng > 0) And (eng <= 4)) Then engFailure(eng) = hasFailure
End Sub

Public Sub SetReverserFailure(ByVal eng As Byte, ByVal hasFailure As Boolean)
    If ((eng > 0) And (eng <= 4)) Then revFailure(eng) = hasFailure
End Sub

Public Function EngineFailure(ByVal eng As Byte) As Boolean
    If ((eng > 0) And (eng <= 4)) Then EngineFailure = engFailure(eng)
End Function

Public Function ReverserFailure(ByVal eng As Byte) As Boolean
    If ((eng > 0) And (eng <= 4)) Then ReverserFailure = revFailure(eng)
End Function

Property Get InstrumentFailureNames() As Variant
    Dim Names As New Collection
    Dim results As Variant
    Dim x As Integer
    
    'Load failure names
    If ADF Then Names.Add "ADF"
    If Altimeter Then Names.Add "Altimeter"
    If Attitude Then Names.Add "Attitude"
    If Airspeed Then Names.Add "Airspeed"
    If COM1 Then Names.Add "COM1"
    If COM2 Then Names.Add "COM2"
    If Compass Then Names.Add "Compass"
    If Fuel Then Names.Add "Fuel"
    If NAV1 Then Names.Add "NAV1"
    If NAV2 Then Names.Add "NAV2"
    If Transponder Then Names.Add "Transponder"
    If VerticalSpeed Then Names.Add "VSI"
    
    ReDim results(Names.Count - 1)
    For x = 1 To Names.Count
        results(x - 1) = Names.Item(x)
    Next
    
    InstrumentFailureNames = results
End Property

Property Get EquipmentFailureNames() As Variant
    Dim Names As New Collection
    Dim results As Variant
    Dim x As Integer
    
    'Load failure names
    If Flaps Then Names.Add "Flaps"
    If Spoilers Then Names.Add "Spoilers"
    If Gear Then Names.Add "Gear"
    If PitotHeat Then Names.Add "PitotHeat"
    For x = 1 To engCount
        If engFailure(x) Then Names.Add "Engine" + CStr(x)
        If revFailure(x) Then Names.Add "Rev" + CStr(x)
    Next
    
    ReDim results(Names.Count - 1)
    For x = 1 To Names.Count
        results(x - 1) = Names.Item(x)
    Next
    
    EquipmentFailureNames = results
End Property

Public Sub DisableFailure(ByVal failureCode As String)
    Dim engine As Integer
    
    'Process engine/reverser failure
    If ((Left(failureCode, 6) = "Engine") Or (Left(failureCode, 3) = "Rev")) Then
        engine = CInt(Right(failureCode, 1))
        failureCode = Left(failureCode, Len(failureCode) - 1)
    End If
    
    Select Case failureCode
        Case "ADF"
            ADF = False
        
        Case "Altimeter"
            Altimeter = False
        
        Case "Attitude"
            Attitude = False
        
        Case "Airspeed"
            Airspeed = False
        
        Case "COM1"
            COM1 = False
        
        Case "COM2"
            COM2 = False
        
        Case "Compass"
            Compass = False
        
        Case "Fuel"
            Fuel = False
        
        Case "NAV1"
            NAV1 = False
        
        Case "NAV2"
            NAV2 = False
        
        Case "Transponder"
            Transponder = False
        
        Case "VSI"
            VerticalSpeed = False
                    
        Case "PitotHeat"
            PitotHeat = False

        Case "Flaps"
            Flaps = False
        
        Case "Spoilers"
            Spoilers = False
        
        Case "Gear"
            Gear = False
        
        Case "Engine"
            engFailure(engine) = False
        
        Case "Rev"
            revFailure(engine) = False
    End Select
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AircraftFuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public AircraftType As String
Public EngineCount As Integer
Public EngineType As String
Public CruiseSpeed As Integer

Public FuelFlow As Long 'Fuel flow/engine/hour
Public BaseFuel As Long
Public TaxiFuel As Long

Public PrimaryTanks As Variant
Public PrimaryPercentage As Integer
Public SecondaryTanks As Variant
Public SecondaryPercentage As Integer
Public OtherTanks As Variant

Property Get CruiseFuel(ByVal Distance As Integer, Optional HeadWind As Integer = 0) As Long
    Dim CruiseTime As Double
    
    CruiseTime = CDbl(Distance) / (CruiseSpeed + HeadWind)
    CruiseFuel = BaseFuel + TaxiFuel + CLng(FuelFlow * CruiseTime * EngineCount)
End Property

Public Sub AddPrimaryTank(ByVal tankName As String)
    ReDim PrimaryTanks(UBound(PrimaryTanks) + 1)
    PrimaryTanks(UBound(PrimaryTanks)) = tankName
End Sub

Public Sub AddSecondaryTank(ByVal tankName As String)
    ReDim SecondaryTanks(UBound(SecondaryTanks) + 1)
    SecondaryTanks(UBound(SecondaryTanks)) = tankName
End Sub

Public Sub AddOtherTank(ByVal tankName As String)
    ReDim OtherTanks(UBound(OtherTanks) + 1)
    OtherTanks(UBound(OtherTanks)) = tankName
End Sub

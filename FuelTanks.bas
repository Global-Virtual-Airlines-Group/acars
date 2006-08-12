Attribute VB_Name = "FuelTanks"
Option Explicit

Public Const CENTER = 0
Public Const LEFT_MAIN = 1
Public Const LEFT_AUX = 2
Public Const LEFT_TIP = 3
Public Const RIGHT_MAIN = 4
Public Const RIGHT_AUX = 5
Public Const RIGHT_TIP = 6

Public Const CENTER2 = 7
Public Const CENTER3 = 8
Public Const EXT1 = 9
Public Const EXT2 = 10

Public Const MAX_TANK = 10
Public Const FUEL_PROFILE_FILE = "FuelProfile.ini"

Public Function AllTankOffsets() As Variant
    AllTankOffsets = Array(&HB74, &HB7C, &HB84, &HB8C, &HB94, &HB9C, &HBA4, &H1244, _
        &H124C, &H1254, &H125C)
End Function

Public Function AllCapacityOffsets() As Variant
    AllCapacityOffsets = Array(&HB78, &HB80, &HB88, &HB90, &HB98, &HBA0, &HBA8, &H1248, _
        &H1250, &H1258, &H1260)
End Function

Public Function TankNames() As Variant
    TankNames = Array("Center", "Left Main", "Left AUX", "Left Tip", "Right Main", _
        "Right AUX", "Right Tip", "Center 2", "Center 3", "External", "External 2")
End Function

Public Function GetTankCode(ByVal tankName As String) As Integer
    Dim x As Integer
    Dim names As Variant
    
    names = TankNames()
    For x = 0 To UBound(names)
        If (UCase(names(x)) = UCase(tankName)) Then
            GetTankCode = x
            Exit Function
        End If
    Next
    
    GetTankCode = -1
End Function

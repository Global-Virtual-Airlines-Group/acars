VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pilot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As String
Public FirstName As String
Public LastName As String
Public Rank As String
Public EquipmentType As String
Public IsBusy As Boolean
Public IsHidden As Boolean

Private Roles As Variant

Public Legs As Integer
Public Hours As Double

Public flightCode As String
Public FlightEQ As String

Public airportD As Airport
Public AirportA As Airport

Public StartTime As Date
Public ClientBuild As Integer
Public RemoteAddress As String
Public RemoteHost As String

Property Get name() As String
    name = FirstName + " " + LastName
End Property

Property Get FlightTotals() As String
    FlightTotals = Format(Legs, "#,##0") + " legs, " + Format(Hours, "##0.0") + " hours"
End Property

Private Sub Class_Initialize()
    Roles = Array()
End Sub

Public Sub AddRole(ByVal RoleName As String)
    If Not HasRole(RoleName) Then
        ReDim Preserve Roles(UBound(Roles) + 1)
        Roles(UBound(Roles)) = RoleName
    End If
End Sub

Public Sub SetRoles(ByVal RoleNames As String)
    Roles = Split(RoleNames, ",")
End Sub

Public Function HasRole(ByVal RoleName As String) As Boolean
    Dim rName As Variant

    If (UBound(Roles) = -1) Then Exit Function
    
    'Loop through the roles
    For Each rName In Roles
        If ((rName = RoleName) Or (rName = "Admin")) Then
            HasRole = True
            Exit Function
        End If
    Next
End Function

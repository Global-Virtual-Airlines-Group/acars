VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Charts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public AirportCode As String
Public AirportName As String

Private Names As Variant
Private IDs As Variant

Public Sub Class_Initialize()
    Names = Array()
    IDs = Array()
End Sub

Public Sub addChart(ByVal name As String, ByVal id As Integer)
    ReDim Preserve Names(UBound(Names) + 1)
    ReDim Preserve IDs(UBound(IDs) + 1)
    Names(UBound(Names)) = name
    IDs(UBound(IDs)) = id
End Sub

Property Get Size()
    Size = UBound(Names) + 1
End Property

Property Get ChartNames() As Variant
    ChartNames = Names
End Property

Property Get ChartIDs() As Variant
    ChartIDs = IDs
End Property

Public Function GetID(ByVal name As String) As Integer
    Dim x As Integer
    
    For x = 0 To UBound(Names)
        If (Names(x) = name) Then
            GetID = x
            Exit Function
        End If
    Next
End Function

Property Get AirportComboEntry()
    AirportComboEntry = AirportName + " (" + AirportCode + ")"
End Property

Public Sub Update()
    Dim x As Integer
    
    frmCharts.cboChart.Clear
    For x = 0 To UBound(Names)
        frmCharts.cboChart.AddItem Names(x)
    Next
End Sub

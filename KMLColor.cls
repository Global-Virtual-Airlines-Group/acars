VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KMLColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Red As Integer
Public Green As Integer
Public Blue As Integer
Public Alpha As Integer

Property Get Text() As String
    Text = ToHex(Alpha) + ToHex(Blue) + ToHex(Green) + ToHex(Red)
End Property

Property Get Adjust(ByVal factor As Double) As KMLColor
    Dim c As New KMLColor
    
    c.Red = Me.Red / factor
    c.Green = Me.Green / factor
    c.Blue = Me.Blue / factor
    c.Alpha = Me.Alpha / factor
    Set Adjust = c
End Property

Public Sub SetColors(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, _
    ByVal a As Integer)
    Red = r
    Green = g
    Blue = b
    Alpha = a
End Sub
    
Private Function ToHex(ByVal color As Integer) As String
    Dim result As String
    
    result = Hex(color)
    If (Len(result) = 1) Then result = "0" + result
    ToHex = result
End Function


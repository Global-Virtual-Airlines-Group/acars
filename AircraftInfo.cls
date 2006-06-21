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

Public HasAfterburner As Boolean
Private acAGL As Integer

Public FSPath As String
Private airP As String
Private airF As String

Public code As String

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

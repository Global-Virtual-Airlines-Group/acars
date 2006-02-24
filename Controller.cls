VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As String
Public NetworkID As String
Public Name As String
Public Frequency As String
Public Rating As String
Public FacilityType As String

Property Get FacilityInfo() As String
    FacilityInfo = ID + " (" + FacilityType + ")"
End Property

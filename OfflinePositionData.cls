VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OfflinePositionData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dataQueue As Variant

Public Sub Class_Initialize()
    dataQueue = Array()
End Sub

Public Sub AddPosition(cPos As PositionData)
    ReDim Preserve dataQueue(UBound(dataQueue) + 1)
    Set dataQueue(UBound(dataQueue)) = cPos
End Sub

Property Get HasData() As Boolean
    HasData = (UBound(dataQueue) > 0)
End Property

Property Get Size() As Integer
    Size = UBound(dataQueue) + 1
End Property

Property Get Queue() As Variant
    Queue = dataQueue
End Property

Public Sub Clear()
    dataQueue = Array()
End Sub

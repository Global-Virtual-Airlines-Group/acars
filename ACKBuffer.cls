VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ACKBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IDs() As Long
Private ReceivedIDs() As Long

Private Sub Class_Initialize()
    ReDim IDs(0)
    ReDim ReceivedIDs(0)
End Sub

Public Sub Queue(ByVal ID As Long)
    Dim x As Integer

    If IsWaitingFor(ID) Then Exit Sub
    
    'Insert into any blank spots
    For x = 0 To UBound(IDs)
        If (IDs(x) = 0) Then
            IDs(x) = ID
            Exit Sub
        End If
    Next
    
    'Add to the end of the array
    ReDim Preserve IDs(UBound(IDs) + 1)
    IDs(UBound(IDs)) = ID
End Sub

Public Sub Receive(ByVal ID As Long)
    Dim x As Integer
    
    If (ID = 0) Or HasReceived(ID) Or Not IsWaitingFor(ID) Then Exit Sub
    
    'Add to the received Queue
    ReDim Preserve ReceivedIDs(UBound(ReceivedIDs) + 1)
    ReceivedIDs(UBound(ReceivedIDs)) = ID
    
    'Remove from the pending Queue
    For x = 0 To UBound(IDs)
        If (IDs(x) = ID) Then
            IDs(x) = 0
            Exit Sub
        End If
    Next
End Sub

Public Function IsWaitingFor(ByVal ID As Long) As Boolean
    Dim x As Integer

    For x = 0 To UBound(IDs)
        If (IDs(x) = ID) Then
            IsWaitingFor = True
            Exit Function
        End If
    Next
End Function

Public Function HasReceived(ByVal ID As Long) As Boolean
    Dim x As Integer
    
    For x = 0 To UBound(ReceivedIDs)
        If (ReceivedIDs(x) = ID) Then
            HasReceived = True
            Exit Function
        End If
    Next
End Function

Public Sub Clear()
    ReDim IDs(0)
    ReDim ReceivedIDs(0)
End Sub

Public Sub Compact()
    Dim x As Integer
    Dim newIDs() As Long
    
    'Build the temporary array
    ReDim newIDs(0)
    For x = 0 To UBound(IDs)
        If (IDs(x) <> 0) Then
            ReDim Preserve newIDs(UBound(newIDs) + 1)
            newIDs(UBound(newIDs)) = IDs(x)
        End If
    Next
    
    'Copy to the main array
    ReDim IDs(UBound(newIDs))
    For x = 0 To UBound(newIDs)
        IDs(x) = newIDs(x)
    Next
End Sub

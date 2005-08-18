VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RequestStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const ACARSVERSION = 1

Private doc As New DOMDocument
Private root As IXMLDOMElement

Private RootDate As Date
Private LastSendTime As Date
Private ReqID As Long

Public Sub Class_Initialize()
    RootDate = DateSerial(2005, 6, 1)
    reset
End Sub

Private Sub reset()
    Dim tmpID As Long
    tmpID = DateDiff("s", RootDate, Now)
    If (tmpID = ReqID) Then
        ReqID = ReqID + 1
    Else
        ReqID = tmpID
    End If

    doc.async = False
    doc.loadXML "<ACARSRequest></ACARSRequest>"
    Set root = doc.documentElement
    
    'Set ACARS version and request ID attributes.
    root.setAttribute "version", ACARSVERSION
    root.setAttribute "id", Hex(ReqID)
End Sub

Property Get LastUse() As Date
    LastUse = LastSendTime
End Property

Property Get RequestID() As Long
    RequestID = ReqID
End Property

Public Sub Queue(cmd As IXMLDOMElement)
    root.appendChild cmd
End Sub

Public Sub Send()
    'Set error handler
    On Error GoTo EH

    'Send the data
    frmMain.wsckMain.SendData "<?xml version=""1.0"" encoding=""UTF-8"" ?>" & doc.XML
    LastSendTime = Now
    If config.ShowDebug Then ShowMessage "Sending Message " + Hex(ReqID), DEBUGTEXTCOLOR
    
ExitSub:
    reset
    Exit Sub

EH:
    MsgBox "The following error occurred: " & err.Description & " (" & err.Number & ")", vbOKOnly Or vbCritical, "tmrRequest.Error"
    Resume ExitSub
End Sub


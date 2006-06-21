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
Public HasData As Boolean

Private doc As New DOMDocument
Private root As IXMLDOMElement

Private RootDate As Date
Private LastSendTime As Date
Private ReqID As Long

Public Sub Class_Initialize()
    RootDate = DateSerial(2005, 6, 1)
    Reset
End Sub

Private Sub Reset()
    Dim tmpID As Long
    Dim pi As IXMLDOMProcessingInstruction
    
    tmpID = DateDiff("s", RootDate, Now)
    If (tmpID <= ReqID) Then
        ReqID = ReqID + 1
    Else
        ReqID = tmpID
    End If
    
    'Create the document
    Set doc = New DOMDocument
    Set pi = doc.createProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")
    doc.insertBefore pi, doc.firstChild

    'Create the root element
    Set root = doc.createNode(NODE_ELEMENT, "ACARSRequest", "")
    root.setAttribute "version", ACARSVERSION
    root.setAttribute "id", Hex(ReqID)
    doc.appendChild root
    HasData = False
End Sub

Property Get LastUse() As Date
    LastUse = LastSendTime
End Property

Property Get RequestID() As Long
    RequestID = ReqID
End Property

Public Sub Queue(cmd As IXMLDOMElement)
    root.appendChild cmd
    HasData = True
End Sub

Public Sub Send()
    On Error GoTo EH
    
    'Send the data
    frmMain.wsckMain.SendData doc.XML
    LastSendTime = Now
    HasData = False
    If config.ShowDebug Then
        ShowMessage "Sending Message " + Hex(ReqID), DEBUGTEXTCOLOR
        ShowDebug doc.XML, XML_OUT_COLOR
    End If
    
ExitSub:
    Reset
    Exit Sub

EH:
    MsgBox "The following error occurred: " & err.Description & " (" & err.Number & ")", vbOKOnly Or vbCritical, "tmrRequest.Error"
    Resume ExitSub
End Sub

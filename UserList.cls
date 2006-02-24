VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pilotInfo As Variant
Private atcInfo As Variant

Public SelectedPilot As String
Public SelectedATC As String

Public Sub Class_Initialize()
    pilotInfo = Array()
    atcInfo = Array()
End Sub

Public Sub AddPilot(newPilot As Pilot)
    Dim p As Pilot
    
    Set p = GetPilot(newPilot.id)
    If Not (p Is Nothing) Then DeletePilot (newPilot.id)

    ReDim Preserve pilotInfo(UBound(pilotInfo) + 1)
    Set pilotInfo(UBound(pilotInfo)) = newPilot
End Sub

Public Sub AddController(newATC As Controller)
    ReDim Preserve atcInfo(UBound(atcInfo) + 1)
    Set atcInfo(UBound(atcInfo)) = newATC
End Sub

Public Sub DeletePilot(ByVal id As String)
    Dim x As Integer
    Dim p As Pilot
    Dim tmpInfo As Variant
    
    tmpInfo = pilotInfo
    pilotInfo = Array()
    For x = 0 To UBound(tmpInfo)
        Set p = tmpInfo(x)
        If (p.id <> UCase(id)) Then AddPilot p
    Next
End Sub

Public Sub ClearPilots()
    pilotInfo = Array()
End Sub

Public Sub ClearATC()
    atcInfo = Array()
End Sub

Public Function GetPilot(ByVal id As String) As Pilot
    Dim x As Integer
    Dim p As Pilot

    id = UCase(id)
    For x = 0 To UBound(pilotInfo)
        Set p = pilotInfo(x)
        If ((p.id = id) Or (UCase(p.Name) = id)) Then
            Set GetPilot = p
            Exit Function
        End If
    Next
End Function

Public Function GetATC(ByVal id As String) As Controller
    Dim x As Integer
    Dim ctr As Controller

    For x = 0 To UBound(atcInfo)
        Set ctr = atcInfo(x)
        If (ctr.id = UCase(id)) Then
            Set GetATC = ctr
            Exit Function
        End If
    Next
End Function

Public Sub UpdatePilotList()
    Dim x As Integer
    Dim p As Pilot
    Dim IsRetained As Boolean
    
    frmMain.lstPilots.enabled = False
    frmMain.lstPilots.Clear
    For x = 0 To UBound(pilotInfo)
        Set p = pilotInfo(x)
        frmMain.lstPilots.AddItem p.id + " (" + p.Name + ")"
        
        'If the pilot didn't log out then keep it selected
        If (p.id = SelectedPilot) Then
            IsRetained = True
            frmMain.lstPilots.ListIndex = x
        End If
    Next
        
    'If we're not selected then clear the info frame
    If Not IsRetained Then frmMain.infoFrame.visible = False
    frmMain.lstPilots.enabled = True
End Sub

Public Sub UpdateATCList()
    Dim x As Integer
    Dim ctr As Object
    Dim IsRetained As Boolean

    frmMain.lstATC.enabled = False
    frmMain.lstATC.Clear
    For x = 0 To UBound(atcInfo)
        Set ctr = atcInfo(x)
        frmMain.lstATC.AddItem ctr.id + " (" + ctr.Name + ") [" + ctr.Frequency + "]"
        
        'If the controller didn't log out then keep it selected
        If (ctr.id = SelectedATC) Then
            IsRetained = True
            frmMain.lstATC.ListIndex = x
        End If
    Next
        
    'If we're not selected then clear the info frame
    frmMain.lstATC.enabled = True
    If Not IsRetained Then
        frmMain.ctrFrame.visible = False
        frmMain.UpdateTuneButtons False, False
    End If
End Sub

Public Function IsRolePresent(ByVal RoleName As String) As Boolean
    Dim x As Integer
    Dim p As Pilot

    For x = 0 To UBound(pilotInfo)
        Set p = pilotInfo(x)
        If p.HasRole(RoleName) Then
            IsRolePresent = True
            Exit Function
        End If
    Next
End Function

Property Get DispatchOnline() As Boolean
    DispatchOnline = IsRolePresent("HR") Or IsRolePresent("Dispatch")
End Property

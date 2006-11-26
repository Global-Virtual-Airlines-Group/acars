Attribute VB_Name = "GaugeInfo"
Option Explicit

Public Const GAUGE_BASE = &H7A20

Public Const ACARS_OFF = 0
Public Const ACARS_ON = 1
Public Const ACARS_CONNECTED = 2

Public Sub GAUGE_SetPhase(ByVal phase As Integer, ByVal IsConnected As Boolean)
    Dim ConnectMode As Integer
    Dim lngResult As Long
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    'Set connection state
    ConnectMode = ACARS_ON
    If IsConnected Then ConnectMode = ACARS_CONNECTED
    Call FSUIPC_Write(GAUGE_BASE, 1, VarPtr(phase), lngResult)
    Call FSUIPC_Write(GAUGE_BASE + 1, 1, VarPtr(ConnectMode), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    If config.ShowDebug Then ShowMessage "Sent flight phase to Gauge", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_SetStatus(ByVal status As Integer)
    Dim lngResult As Long
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    Call FSUIPC_Write(GAUGE_BASE + 1, 1, VarPtr(status), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    If config.ShowDebug Then ShowMessage "Notified Gauge of ACARS Status", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_SetInfo(fInfo As FlightData, PilotID As String)
    Dim pID As String
    Dim fID As String
    Dim lngResult As Long
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    'Null-terminate strings passed to FSUIPC
    pID = PilotID & Chr(0)
    fID = fInfo.flightCode & Chr(0)
    
    'Update connection status
    If config.ACARSConnected Then
        Call FSUIPC_Write(GAUGE_BASE + 1, 1, VarPtr(ACARS_CONNECTED), lngResult)
    Else
        Call FSUIPC_Write(GAUGE_BASE + 1, 1, VarPtr(ACARS_ON), lngResult)
    End If
    
    'Write flight information
    Call FSUIPC_Write(GAUGE_BASE + 2, 4, VarPtr(fInfo.FlightID), lngResult)
    Call FSUIPC_WriteS(GAUGE_BASE + 6, Len(pID), pID, lngResult)
    Call FSUIPC_WriteS(GAUGE_BASE + 16, Len(fID), fID, lngResult)
    Call FSUIPC_Write(GAUGE_BASE + 26, 1, VarPtr(fInfo.FlightLeg), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting Gauge Flight Information", ACARSERRORCOLOR
    ElseIf config.ShowDebug Then
        ShowMessage "Sent Flight Information to Gauge", DEBUGGAUCOLOR
    End If
End Sub

Public Sub GAUGE_SetChat(Optional IsPrivate As Boolean = False)
    Dim MsgFlag As Integer
    Dim lngResult As Long
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    MsgFlag = 1
    If IsPrivate Then MsgFlag = 2
    Call FSUIPC_Write(GAUGE_BASE + 27, 1, VarPtr(MsgFlag), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    If config.ShowDebug Then ShowMessage "Notified Gauge of chat message", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_ClearChat()
    Dim MsgFlag As Integer
    Dim lngResult As Long
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    MsgFlag = 0
    Call FSUIPC_Write(GAUGE_BASE + 27, 1, VarPtr(MsgFlag), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    If config.ShowDebug Then ShowMessage "Cleared Gauge chat message indicator", DEBUGGAUCOLOR
End Sub

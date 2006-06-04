Attribute VB_Name = "GaugeInfo"
Option Explicit

Private Const FSUIPC_BASE = &H8100

Private lngResult As Long

Public Sub GAUGE_SetPhase(phase As Integer, IsConnected As Boolean)
    Dim ACARSConnected As Integer
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    If ACARSConnected Then IsConnected = 1
    Call FSUIPC_Write(FSUIPC_BASE, 1, VarPtr(phase), lngResult)
    Call FSUIPC_Write(FSUIPC_BASE + 1, 1, VarPtr(IsConnected), lngResult)
    
    'Log debug message
    ShowMessage "Sent flight phase to Gauge", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_Disconnect()
    Dim IsConnected As Integer
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    IsConnected = 0
    Call FSUIPC_Write(FSUIPC_BASE + 1, 1, VarPtr(IsConnected), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    ShowMessage "Notified Gauge of ACARS Disconnect", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_SetInfo(fInfo As FlightData, PilotID As String)
    Dim pID As String
    Dim fID As String
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    pID = PilotID & Chr(0)
    fID = fInfo.flightCode & Chr(0)
    
    Call FSUIPC_Write(FSUIPC_BASE + 2, 4, VarPtr(fInfo.FlightID), lngResult)
    Call FSUIPC_WriteS(FSUIPC_BASE + 6, Len(pID), pID, lngResult)
    Call FSUIPC_WriteS(FSUIPC_BASE + 16, Len(fID), fID, lngResult)
    Call FSUIPC_Write(FSUIPC_BASE + 26, 1, VarPtr(fInfo.FlightLeg), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    ShowMessage "Sent Flight Information to Gauge", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_SetChat(Optional IsPrivate As Boolean = False)
    Dim MsgFlag As Integer
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    MsgFlag = 1
    If IsPrivate Then MsgFlag = 2
    Call FSUIPC_Write(FSUIPC_BASE + 36, 1, VarPtr(MsgFlag), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    ShowMessage "Notified Gauge of chat message", DEBUGGAUCOLOR
End Sub

Public Sub GAUGE_ClearChat()
    Dim MsgFlag As Integer
    
    'Do nothing if gauge support disabled
    If Not config.GaugeSupport Or Not config.FSUIPCConnected Then Exit Sub
    
    MsgFlag = 0
    Call FSUIPC_Write(FSUIPC_BASE + 36, 1, VarPtr(MsgFlag), lngResult)
    Call FSUIPC_Process(lngResult)
    
    'Log debug message
    ShowMessage "Cleared Gauge chat message indicator", DEBUGGAUCOLOR
End Sub

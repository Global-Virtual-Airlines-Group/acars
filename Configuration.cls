VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Registry settings key
Private Const REG_ROOT_KEY = "SOFTWARE\Delta Virtual Airlines"
Private Const REG_SETTINGS_KEY = "SOFTWARE\Delta Virtual Airlines\ACARS"

'Airport configuration file
Private Const AIRPORT_DATA = "airports.cfg"
Private Const EQ_DATA = "eqTypes.cfg"

'Server connection constants.
Private Const DEFAULT_ACARS_HOST = "golgotha.deltava.org"
Private Const DEFAULT_ACARS_PORT = 15527
Private Const DEFAULT_POS_INTERVAL = 30
Private Const DEFAULT_PING_INTERVAL = 45

'ACARS Server data
Private ACARS_Host As String
Private ACARS_Port As Integer

Private eqTypes As Variant
Private aNames As Variant
Private aCodes As Variant

'Flight state data
Public FSUIPCConnected As Boolean
Public ACARSConnected As Boolean
Public SeenHELO As Boolean

'Menu options
Public FlyOffline As Boolean
Public PlaySound As Boolean
Public SavePassword As Boolean
Public ShowTimestamps As Boolean
Public ShowDebug As Boolean

'Update intervals
Public PingInterval As Integer
Public PositionInterval As Integer

'SB3 Support info
Public SB3Support As Boolean
Public SB3Installed As Boolean
Public SB3Connected As Boolean

'IVAp Support info
Public IVAPSupport As Boolean
Public IVAPInstalled As Boolean
Public IVAPConnected As Boolean

Public PrivateVoiceURL As String

Public Sub Class_Initialize()
    eqTypes = Array("A310", "A318", "A319", "A320", "A321", "A330", "A340", "ATR-72", "B707-320", "B727-200", "B737-200", "B737-300", "B737-800", "B747-100", "B747-200", "B747-400", "B757-200", "B767-200", "B767-300", "B767-400", "B777-200", "B777-300", "Caravelle", "Comet", "Concorde", "Constellation", "CRJ-200", "CRJ-700", "CV-880", "DC-10-10", "DC-3", "DC-6", "DC-8-11", "DC-8-33", "DC-8-51", "DC-8-61", "DC-8-71", "DC-9-14", "DC-9-32", "DO-328", "EMB-120", "ERJ-135", "ERJ-145", "L-1011", "MD-11", "MD-88", "MD-90")
    aNames = Array()
    aCodes = Array()

    'Load Global Settings
    ACARS_Host = RegReadString(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ACARSServerHost", DEFAULT_ACARS_HOST)
    ACARS_Port = RegReadDWORD(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ACARSServerPort", DEFAULT_ACARS_PORT)

    'Load Local Settings
    frmMain.txtPilotID.Text = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Pilot ID", "")
    frmMain.txtPassword.Text = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Password", "")
    frmMain.txtFlightNumber.Text = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightNumber", "")
    info.EquipmentType = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Equipment", "")
    info.CruiseAltitude = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Altitude", "")
    info.Route = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Route", "")
    info.Remarks = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Remarks", "")
    info.AirportD = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "AirportD", "")
    info.AirportA = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "AirportA", "")
    info.AirportL = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "AirportL", "")
    
    'Load the Pilot ID from alternate source if not found
    If (frmMain.txtPilotID.Text = "") Then
        frmMain.txtPilotID.Text = RegReadString(HKEY_CURRENT_USER, REG_ROOT_KEY, "Pilot ID", "")
    End If

    'Set intervals in seconds
    PositionInterval = 30
    PingInterval = 60
    
    'Check if SB3 is installed
    SB3Installed = SB3InstallCheck()

    'Load button states
    SavePassword = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "SavePassword", True)
    PlaySound = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "PlaySoundForChat", True)
    FlyOffline = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "FlyOffline", False)
    ShowTimestamps = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowTimestamps", True)
    ShowDebug = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowDebug", False)
    SB3Support = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "SB3Support", False) And SB3Installed
    IVAPSupport = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "IVAPSupport", False) And False
End Sub

Public Sub UpdateSettingsMenu()
    frmMain.mnuOptionsFlyOffline.Checked = FlyOffline
    frmMain.mnuOptionsPlaySoundForChat.Checked = PlaySound
    frmMain.mnuOptionsSavePassword.Checked = SavePassword
    frmMain.mnuOptionsShowTimestamps.Checked = ShowTimestamps
    frmMain.mnuOptionsShowDebugMessages.Checked = ShowDebug
    frmMain.mnuOptionsSB3Support.Enabled = SB3Installed
    frmMain.mnuOptionsSB3Support.Checked = SB3Support
    frmMain.mnuOptionsIVAPSupport.Checked = IVAPSupport
    
    'Update SB3 support
    frmMain.mnuSaveFlightPlan.Enabled = SB3Support
End Sub

Public Sub UpdateFlightInfo()
    frmMain.txtCruiseAlt.Text = info.CruiseAltitude
    frmMain.txtRemarks.Text = info.Remarks
    frmMain.txtRoute.Text = info.Route
    frmMain.cboAirportD.ListIndex = GetAirport(info.AirportD) + 1
    frmMain.cboAirportA.ListIndex = GetAirport(info.AirportA) + 1
    frmMain.cboAirportL.ListIndex = GetAirport(info.AirportL) + 1
End Sub

Public Sub Save()
    'Save Flight Data to registry
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Pilot ID", frmMain.txtPilotID.Text
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Password", IIf(SavePassword, frmMain.txtPassword.Text, "")
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightNumber", frmMain.txtFlightNumber.Text
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Equipment", info.EquipmentType
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Altitude", info.CruiseAltitude
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "AirportD", info.AirportD
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "AirportA", info.AirportA
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "AirportL", info.AirportL
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Route", info.Route
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Remarks", info.Remarks

    'Save settings
    RegWriteBoolean HKEY_CURRENT_USER, REG_SETTINGS_KEY, "SavePassword", SavePassword
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "PlaySoundForChat", PlaySound
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "FlyOffline", FlyOffline
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowTimestamps", ShowTimestamps
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowDebug", ShowDebug
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "SB3Support", SB3Support
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "IVAPSupport", IVAPSupport
End Sub

Property Get ACARSHost() As String
    ACARSHost = ACARS_Host
End Property

Property Get ACARSPort() As Integer
    ACARSPort = ACARS_Port
End Property

Property Get AirportNames() As Variant
    AirportNames = aNames
End Property

Property Get AirportCodes() As Variant
    AirportCodes = aCodes
End Property

Property Get EquipmentTypes() As Variant
    EquipmentTypes = eqTypes
End Property

Public Sub AddAirport(ByVal name As String, ByVal icao As String)
    ReDim Preserve aNames(UBound(aNames) + 1)
    ReDim Preserve aCodes(UBound(aCodes) + 1)
    aNames(UBound(aNames)) = name
    aCodes(UBound(aCodes)) = icao
End Sub

Public Sub AddEquipment(ByVal eqType As String)
    ReDim Preserve eqTypes(UBound(eqTypes) + 1)
    eqTypes(UBound(eqTypes)) = eqType
End Sub

Public Sub ClearEquipment()
    eqTypes = Array()
End Sub

Public Function IsConfigUpToDate() As Boolean
    Dim apDate As Date
    Dim eqDate As Date
    Dim age As Integer

    'Calculate average age of the data files
    apDate = FileDateTime(App.Path + "\" + AIRPORT_DATA)
    eqDate = FileDateTime(App.Path + "\" + EQ_DATA)
    age = (DateDiff("d", apDate, Now) + DateDiff("d", eqDate, Now)) / 2
    
    IsConfigUpToDate = (age <= 21)
End Function

Public Sub SaveEquipment()
    Dim fNum As Integer
    Dim x As Integer
    
    On Error GoTo FatalError
    
    fNum = FreeFile()
    Open App.Path + "\" + EQ_DATA For Output As #fNum
    For x = 0 To UBound(eqTypes)
        Print #fNum, eqTypes(x)
    Next

    Close #fNum
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot save Equipment Types!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub

End Sub

Public Sub LoadEquipment()
    Dim fNum As Integer
    Dim rawData As Variant
    Dim eq As Variant
    
    On Error GoTo FatalError
    
    fNum = FreeFile()
    Open App.Path + "\" + EQ_DATA For Input As #fNum
    rawData = Split(Input(LOF(fNum), fNum), vbCrLf)
    Close #fNum
    
    ClearEquipment
    For Each eq In rawData
        AddEquipment eq
    Next
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot load Equipment Types!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub
    
End Sub

Public Sub ClearAirports()
    aNames = Array()
    aCodes = Array()
End Sub

Public Sub SaveAirports()
    Dim fNum As Integer
    Dim x As Integer
    
    On Error GoTo FatalError
    
    fNum = FreeFile()
    Open App.Path + "\" + AIRPORT_DATA For Output As #fNum
    For x = 0 To UBound(aNames)
        Print #fNum, aNames(x) + "," + aCodes(x)
    Next
    
    Close #fNum
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot save Airport data!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub

End Sub

Public Sub LoadAirports()
    Dim fNum As Integer
    Dim rawData As Variant
    Dim apData As Variant
    Dim rawAirport As Variant
    
    On Error GoTo FatalError

    fNum = FreeFile()
    Open App.Path + "\" + AIRPORT_DATA For Input As #fNum
    rawData = Split(Input(LOF(fNum), fNum), vbCrLf)
    Close #fNum
    
    For Each rawAirport In rawData
        apData = Split(rawAirport, ",")
        If (UBound(apData) > 0) Then AddAirport apData(0), apData(1)
    Next
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot load Airport data!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub

End Sub

Public Function GetAirport(ByVal code As String) As Integer
    Dim x As Integer
    
    For x = 0 To UBound(aCodes)
        If (code = aCodes(x)) Then
            GetAirport = x
            Exit Function
        End If
    Next
    
    GetAirport = 0
End Function

Public Sub SaveFlightInfo(FlightID As Long, Optional startTime As Date)
    If (FlightID > 0) Then
        RegWriteDWORD HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightID", FlightID
        RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightStartTime", CStr(startTime)
    Else
        RegDelValue HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightID"
        RegDelValue HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightStartTime"
    End If
End Sub

Public Function LoadFlightInfo() As FlightData
    Dim fInfo As New FlightData
    
    fInfo.FlightID = RegReadDWORD(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightID", 0)
    fInfo.startTime = CDate(RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightStartTime", CStr(Now)))
    Set LoadFlightInfo = fInfo
End Function

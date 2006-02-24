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
Private Const INT_FORMAT_KEY = "Control Panel\International"

'Airport configuration file
Private Const AIRPORT_DATA = "airports.cfg"
Private Const EQ_DATA = "eqTypes.cfg"

'Server connection constants.
Private Const DEFAULT_ACARS_HOST = "acars.deltava.org"
Private Const DEFAULT_ACARS_PORT = 15527
Private Const DEFAULT_POS_INTERVAL = 30
Private Const DEFAULT_PING_INTERVAL = 45

'ACARS Server data
Private ACARS_Host As String
Private ACARS_Port As Integer

Private eqTypes As Variant
Private airportInfo As Variant

'Flight state data
Public FSUIPCConnected As Boolean
Public ACARSConnected As Boolean
Public SeenHELO As Boolean

'Decimal/date/time separators
Public DecimalSeparator As String
Public DateSeparator As String
Public TimeSeparator As String

'Menu options
Public FlyOffline As Boolean
Public PlaySound As Boolean
Public SavePassword As Boolean
Public ShowTimestamps As Boolean
Public ShowDebug As Boolean
Public ShowPilotNames As Boolean
Public HideWhenMinimized As Boolean
Public HideMessagesWhenBusy As Boolean
Public SterileCockpit As Boolean
Public GaugeSupport As Boolean
Public FreezeWindow As Boolean

'Security roles
Private Roles As Variant
Public IsUnrestricted As Boolean
Public NoMessages As Boolean

'Update intervals
Public PingInterval As Integer
Public PositionInterval As Integer

'WideFS support - this overrides SB3/IVAp installation checks
Public WideFSInstalled As Boolean

'SB3 Support info
Public SB3Support As Boolean
Public SB3Installed As Boolean
Public SB3Connected As Boolean

'TeamSpeak support info
Public TS2Support As Boolean

Public PrivateVoiceURL As String
Public MsgReceived As Boolean
Public Busy As Boolean

Private FS9Data As String
Private ForceCfgUpdate As Boolean

'Win32 API constants for user shell folders
Private Const CSIDL_MYDOCUMENTS = &H5&
Private Const CSIDL_APPDATA = &H1A&
Private Const CSIDL_LOCAL_APPDATA = &H1C&

'Win32 API functions to get My Docs folder
'Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" _
  (ByVal hWnd As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, _
  ByVal lpszPath As String) As Long

Public Sub Class_Initialize()
    'Load the decimal/date separator
    DecimalSeparator = RegReadString(HKEY_CURRENT_USER, INT_FORMAT_KEY, "sDecimal", ".")
    DateSeparator = RegReadString(HKEY_CURRENT_USER, INT_FORMAT_KEY, "sDate", "/")
    TimeSeparator = RegReadString(HKEY_CURRENT_USER, INT_FORMAT_KEY, "sTime", ":")

    eqTypes = Array("A310", "A318", "A319", "A320", "A321", "A330-200", "A340-300", "ATR-72", "B707-320", "B727-200", "B737-200", "B737-300", "B737-800", "B747-100", "B747-200", "B747-400", "B757-200", "B767-200", "B767-300", "B767-400", "B777-200", "B777-300", "Caravelle", "Comet", "Concorde", "Constellation", "CRJ-200", "CRJ-700", "CV-880", "DC-10-10", "DC-3", "DC-6", "DC-8-11", "DC-8-33", "DC-8-51", "DC-8-61", "DC-8-71", "DC-9-14", "DC-9-32", "DO-328", "EMB-120", "ERJ-135", "ERJ-145", "L-1011", "MD-11", "MD-88", "MD-90")
    airportInfo = Array()
    Roles = Array()

    'Load Global Settings
    ACARS_Host = RegReadString(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ACARSServerHost", DEFAULT_ACARS_HOST)
    ACARS_Port = RegReadDWORD(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ACARSServerPort", DEFAULT_ACARS_PORT)
    SeenHELO = False

    'Load Local Settings
    frmMain.txtPilotID.Text = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Pilot ID", "")
    frmMain.txtPassword.Text = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Password", "")
    
    'Load the Pilot ID from alternate source if not found
    If (frmMain.txtPilotID.Text = "") Then
        frmMain.txtPilotID.Text = RegReadString(HKEY_CURRENT_USER, REG_ROOT_KEY, "Pilot ID", "")
    End If

    'Set intervals in seconds
    PositionInterval = 30
    PingInterval = 60
    
    'Check if SB3 is installed
    SB3Installed = SB3InstallCheck()
    
    'Check for config update flag
    ForceCfgUpdate = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "UpdateConfig", False)

    'Load button states
    SavePassword = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "SavePassword", True)
    ShowPilotNames = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "ShowUserNames", False)
    HideMessagesWhenBusy = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "HideMessagesWhenBusy", False)
    HideWhenMinimized = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "HideOnMinimize", False)
    SterileCockpit = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "SterileCockpit", False)
    PlaySound = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "PlaySoundForChat", True)
    FlyOffline = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "FlyOffline", False)
    GaugeSupport = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "GaugeIntegration", False)
    FreezeWindow = RegReadBoolean(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FreezeWindow", False)
    ShowTimestamps = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowTimestamps", True)
    ShowDebug = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowDebug", False)
    WideFSInstalled = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "WideFS", False)
    TS2Support = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "TS2Support", False) _
        And TS2Installed
    SB3Support = RegReadBoolean(HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "SB3Support", False) _
        And (SB3Installed Or WideFSInstalled)
End Sub

Property Get IsFS9() As Boolean
    IsFS9 = (FS9Data <> "")
End Property

Property Get FS9Files() As String
    FS9Files = FS9Data
End Property

Property Let IsFS9(fs9 As Boolean)
    If Not fs9 Then
        FS9Data = ""
        Exit Property
    End If
    
    'Load the My Documents folder
    Dim buf As String
    buf = Space$(260)
    If (SHGetFolderPath(frmMain.hWnd, CSIDL_MYDOCUMENTS, 0&, -1&, buf) = 0) Then
        FS9Data = Left(buf, InStr(buf, vbNullChar) - 1) + "\Flight Simulator Files"
        If ShowDebug Then ShowMessage "Detected FS9 data folder " + FS9Data, DEBUGTEXTCOLOR
    Else
        ShowMessage "Cannot find My Documents folder - FS9 Data persistence disabled", ACARSERRORCOLOR
    End If
End Property

Public Sub UpdateSettingsMenu()
    With frmMain
        .mnuOptionsFlyOffline.Checked = FlyOffline
        .mnuCOMPlaySoundForChat.Checked = PlaySound
        .mnuOptionsSavePassword.Checked = SavePassword
        .mnuCOMShowTimestamps.Checked = ShowTimestamps
        .mnuOptionsShowPilotNames.Checked = ShowPilotNames
        .mnuOptionsGaugeIntegration.Checked = GaugeSupport
        .mnuOptionsShowDebugMessages.Checked = ShowDebug
        .mnuOptionsHide.Checked = HideWhenMinimized
        .mnuCOMHideBusy.Checked = HideMessagesWhenBusy
        .mnuCOMFreezeWindow.Checked = FreezeWindow
        .mnuCOMSterileCockpit.Checked = SterileCockpit
        .mnuOptionsSB3Support.enabled = SB3Installed Or WideFSInstalled
        .mnuOptionsSB3Support.Checked = SB3Support
        .mnuOptionsTS2Support.enabled = TS2Installed And False
        .mnuOptionsTS2Support.Checked = TS2Support
        .mnuOptionsWideFS.Checked = WideFSInstalled
        
        'Update SB3 support
        .mnuSaveFlightPlan.enabled = SB3Support
        .SSTab1.TabVisible(3) = ShowDebug
    End With
End Sub

Public Sub UpdateFlightInfo()
    frmMain.txtFlightNumber.Text = info.FlightNumber
    frmMain.txtCruiseAlt.Text = info.CruiseAltitude
    frmMain.txtRemarks.Text = info.Remarks
    frmMain.txtRoute.Text = info.Route
    frmMain.cboNetwork.Text = info.Network
    frmMain.cboAirportD.ListIndex = GetAirportOfs(info.airportD) + 1
    frmMain.cboAirportA.ListIndex = GetAirportOfs(info.AirportA) + 1
    frmMain.cboAirportL.ListIndex = GetAirportOfs(info.AirportL) + 1
    If info.CheckRide Then frmMain.chkCheckRide.value = 1
End Sub

Public Sub Save()
    'Save Flight Data to registry
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Pilot ID", frmMain.txtPilotID.Text
    RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "Password", IIf(SavePassword, frmMain.txtPassword.Text, "")

    'Save settings
    RegWriteBoolean HKEY_CURRENT_USER, REG_SETTINGS_KEY, "SavePassword", SavePassword
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "PlaySoundForChat", PlaySound
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "GaugeIntegration", GaugeSupport
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "FlyOffline", FlyOffline
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowTimestamps", ShowTimestamps
    RegWriteBoolean HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FreezeWindow", FreezeWindow
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "ShowDebug", ShowDebug
    RegWriteBoolean HKEY_CURRENT_USER, REG_SETTINGS_KEY, "HideOnMinimize", HideWhenMinimized
    RegWriteBoolean HKEY_CURRENT_USER, REG_SETTINGS_KEY, "HideMessagesWhenBusy", HideMessagesWhenBusy
    RegWriteBoolean HKEY_CURRENT_USER, REG_SETTINGS_KEY, "SterileCockpit", SterileCockpit
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "WideFS", WideFSInstalled
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "SB3Support", SB3Support
    RegWriteBoolean HKEY_LOCAL_MACHINE, REG_SETTINGS_KEY, "TS2Support", TS2Support
End Sub

Property Get ACARSHost() As String
    ACARSHost = ACARS_Host
End Property

Property Get ACARSPort() As Integer
    ACARSPort = ACARS_Port
End Property

Property Get Airports() As Variant
    Airports = airportInfo
End Property

Property Get AirportNames() As Variant
    Dim x As Integer
    Dim a As Airport
    Dim result As Variant
    
    ReDim result(UBound(airportInfo))
    For x = 0 To UBound(airportInfo)
        Set a = airportInfo(x)
        result(x) = a.Name + " (" + a.ICAO + ")"
    Next
        
    AirportNames = result
End Property

Property Get EquipmentTypes() As Variant
    EquipmentTypes = eqTypes
End Property

Public Sub AddAirport(ap As Airport)
    ReDim Preserve airportInfo(UBound(airportInfo) + 1)
    Set airportInfo(UBound(airportInfo)) = ap
End Sub

Public Function GetAirportByOfs(ofs As Integer) As Airport
    If (ofs >= 0) Then Set GetAirportByOfs = airportInfo(ofs)
End Function

Public Sub AddEquipment(ByVal eqType As String)
    ReDim Preserve eqTypes(UBound(eqTypes) + 1)
    eqTypes(UBound(eqTypes)) = eqType
End Sub

Public Sub ClearEquipment()
    eqTypes = Array()
End Sub

Public Sub ClearRoles()
    Roles = Array()
End Sub

Public Sub AddRole(ByVal RoleName As String)
    If Not HasRole(RoleName) Then
        ReDim Preserve Roles(UBound(Roles) + 1)
        Roles(UBound(Roles)) = RoleName
    End If
End Sub

Public Function HasRole(ByVal RoleName As String) As Boolean
    Dim rName As Variant

    If (UBound(Roles) = -1) Then Exit Function
    
    'Loop through the roles
    For Each rName In Roles
        If ((rName = RoleName) Or (rName = "Admin")) Then
            HasRole = True
            Exit Function
        End If
    Next
End Function

Public Function IsConfigUpToDate() As Boolean
    Dim apDate As Date
    Dim eqDate As Date
    Dim age As Integer
    
    'Check for override flag
    If ForceCfgUpdate Then
        IsConfigUpToDate = True
        Exit Function
    End If

    'Set error handler
    On Error GoTo FatalError

    'Calculate average age of the data files
    apDate = FileDateTime(App.path + "\" + AIRPORT_DATA)
    eqDate = FileDateTime(App.path + "\" + EQ_DATA)
    age = (DateDiff("d", apDate, Now) + DateDiff("d", eqDate, Now)) / 2
    
    IsConfigUpToDate = (age <= 21)
    
ExitFunc:
    Exit Function
    
FatalError:
    Resume ExitFunc
    
End Function

Public Sub SaveEquipment()
    Dim fNum As Integer
    Dim x As Integer
    
    On Error GoTo FatalError
    
    fNum = FreeFile()
    Open App.path + "\" + EQ_DATA For Output As #fNum
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
    Open App.path + "\" + EQ_DATA For Input As #fNum
    rawData = Split(Input(LOF(fNum), fNum), vbCrLf)
    Close #fNum
    
    ClearEquipment
    For Each eq In rawData
        If (Len(eq) > 1) Then AddEquipment eq
    Next
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot load Equipment Types!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub
    
End Sub

Public Sub ClearAirports()
    airportInfo = Array()
End Sub

Public Sub SaveAirports()
    Dim fNum As Integer
    Dim x As Integer
    Dim a As Airport
    
    Dim lat As String
    Dim lng As String
    
    'Set critical error handler
    On Error GoTo FatalError
    
    'Save the airports
    fNum = FreeFile()
    Open App.path + "\" + AIRPORT_DATA For Output As #fNum
    For x = 0 To UBound(airportInfo)
        Set a = airportInfo(x)
        lat = Replace(Format(a.Latitude, "##0.00000"), DecimalSeparator, ".")
        lng = Replace(Format(a.Longitude, "##0.00000"), DecimalSeparator, ".")
        Print #fNum, a.Name + "," + a.IATA + "," + a.ICAO + "," + lat + "," + lng
    Next
    
    Close #fNum
    
    'Clear the force update flag if present
    ForceCfgUpdate = False
    RegDelValue HKEY_CURRENT_USER, REG_SETTINGS_KEY, "UpdateConfig"
    
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
    Dim a As Airport
    
    On Error GoTo FatalError

    fNum = FreeFile()
    Open App.path + "\" + AIRPORT_DATA For Input As #fNum
    rawData = Split(Input(LOF(fNum), fNum), vbCrLf)
    Close #fNum
    
    For Each rawAirport In rawData
        apData = Split(rawAirport, ",")
        If (UBound(apData) > 0) Then
            Set a = New Airport
            a.Name = apData(0)
            a.IATA = apData(1)
            If (UBound(apData) > 1) Then
                a.ICAO = apData(2)
                a.Latitude = CDbl(Replace(apData(3), ".", DecimalSeparator))
                a.Longitude = CDbl(Replace(apData(4), ".", DecimalSeparator))
            End If
            
            AddAirport a
        End If
    Next
    
ExitSub:
    Exit Sub
    
FatalError:
    MsgBox "Cannot load Airport data!", vbOKOnly + vbCritical, "I/O Error"
    Resume ExitSub

End Sub

Public Function GetAirport(ByVal code As String) As Airport
    Dim x As Integer
    Dim a As Airport
    
    For x = 0 To UBound(airportInfo)
        Set a = airportInfo(x)
        If ((code = a.IATA) Or (code = a.ICAO)) Then
            Set GetAirport = a
            Exit Function
        End If
    Next
End Function

Private Function GetAirportOfs(a As Airport) As Integer
    Dim x As Integer
    Dim ap As Airport
    
    GetAirportOfs = -1
    If (a Is Nothing) Then Exit Function
    For x = 0 To UBound(airportInfo)
        Set ap = airportInfo(x)
        If (a.IATA = ap.IATA) Then
            GetAirportOfs = x
            Exit Function
        End If
    Next
End Function

Public Sub SaveFlightCode(flightCode As String)
    If IsFS9 And (flightCode <> "") Then
        RegWriteString HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightID", flightCode
    Else
        RegDelValue HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightID"
    End If
End Sub

Public Function OldFlightID() As String
    OldFlightID = RegReadString(HKEY_CURRENT_USER, REG_SETTINGS_KEY, "FlightID", "")
End Function

Public Function GetClosestAirport(ByVal lat As Double, ByVal lng As Double) As Airport
    Dim result As Airport, a As Airport
    Dim distance As Integer
    Dim x As Integer

    distance = 15000
    For x = 0 To UBound(airportInfo)
        Set a = airportInfo(x)
        If (a.DistanceTo(lat, lng) < distance) Then
            Set result = a
            distance = a.DistanceTo(lat, lng)
        End If
    Next
    
    Set GetClosestAirport = result
End Function

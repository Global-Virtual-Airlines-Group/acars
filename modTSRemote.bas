Attribute VB_Name = "TSRemote"
'#### Teamspeak Remote Control Modul for Visual Basic 6
'#### For TSRemote.dll of Client Version: 2.0.32.60
'####
'#### (C)2004 René Schädlich (from Germany) - MrCPU@gmx.net
'####
'#### Sorry for my english, its not so good. :P
'####
'#### Subs / Functions:
'####   - TSGetLastError
'####   - TSConnect(Server, Nick, Login, Password, Channel, Password)
'####   - TSDisconnect
'####   - TSQuit
'####   - TSGetClientVersion
'####   - TSSwitchChannelByID(ID, Password)
'####   - TSMuteSound(Mute)
'####   - TSSetAway(Away)
'####   - TSSendMessageToChannel(ID, Text)
'####   - TSKickUserFromChannel(ID, Reason)
'####   - TSKickUserFromServer(ID, Reason)
'####   - TSGetServerInfo
'####   - TSGetPlayerInfoByID(ID)
'####   - TSGetChannelList
'####   - TSGetUserList
'####   - TSGetSpeakers
'####
'#### Additional Subs / Functions:
'####   - TSGetClientPath
'####   - TSGetClientWindowAndStatusHandle
'####   - TSGetStatusTextLen
'####   - TSGetStatusText
'####
'#### Last Update: 23. Mar. 2004
'####
'#### Please do not remove the copyright!
'#### Just add some lines, if you want, when you change something.

'###########################
'#### Teamspeak Remote Types
'###########################
Private Type tsrVersion
 Major As Long
 Minor As Long
 Release As Long
 Build As Long
End Type

Public Type tsrServerInfo
 ServerName As String
 WelcomeMessage As String
 ServerVersion As String
 ServerPlatform As String
 ServerIp As String
 ServerHost As String
 ServerType As String
 ServerMaxUsers As Long
 SupportedCodecs As String
 ChannelCount As Long
 PlayerCount As Long
End Type

Private Type tsrServerInfoGet
 ServerName(29) As Byte
 WelcomeMessage(255) As Byte
 ServerVMajor(3) As Byte
 ServerVMinor(3) As Byte
 ServerVRelease(3) As Byte
 ServerVBuild(3) As Byte
 ServerPlatform(29) As Byte
 ServerIp(29) As Byte
 ServerHost(99) As Byte
 ServerType(3) As Byte
 ServerMaxUsers(3) As Byte
 SupportedCodecs(3) As Byte
 ChannelCount(3) As Byte
 PlayerCount(3) As Byte
End Type

Public Type tsrPlayerInfo
 PlayerID As Long
 ChannelID As Long
 NickName As String
 PlayerChannelPrivileges As String
 PlayerPrivileges As String
 PlayerFlags As String
End Type

Private Type tsrPlayerInfoGet
 PlayerID(3) As Byte
 ChannelID(3) As Byte
 NickName(29) As Byte
 PlayerChannelPrivileges(3) As Byte
 PlayerPrivileges(3) As Byte
 PlayerFlags(3) As Byte
End Type

Public Type tsrChannelInfo
 ChannelID As Long
 ChannelParentID As Long
 PlayerCountInChannel As Long
 ChannelFlags As String
 Codec As String
 ChannelName As String
End Type

Private Type tsrChannelInfoGet
 ChannelID(3) As Byte
 ChannelParentID(3) As Byte
 PlayerCountInChannel(3) As Byte
 ChannelFlags(3) As Byte
 Codec(3) As Byte
 ChannelName(29) As Byte
End Type

Private Type tsrUserInfo
 Player As tsrPlayerInfoGet
 Channel As tsrChannelInfoGet
 ParentChannel As tsrChannelInfoGet
End Type

'#############
'#### Own Enum
'#############
Public Enum tsStatus
 [TSUNKNOWN] = -1
 [ClientNotRunning] = 0
 [NotConnected] = 1
 [Connected] = 2
End Enum

'###############################
'#### Teamspeak Remote Functions
'###############################
Private Declare Function tsrGetLastError Lib "TSRemote" (ErrorMessage As String, EMLen As Long) As Long
Private Declare Function tsrConnect Lib "TSRemote" (Url As String) As Long
Private Declare Function tsrDisconnect Lib "TSRemote" () As Long
Private Declare Function tsrQuit Lib "TSRemote" () As Long
Private Declare Function tsrGetVersion Lib "TSRemote" () As tsrVersion
Private Declare Function tsrGetServerInfo Lib "TSRemote" () As tsrServerInfoGet
Private Declare Function tsrGetPlayerInfoByID Lib "TSRemote" (ID As Long, PlayerInfo As tsrPlayerInfoGet) As Long
Private Declare Function tsrGetChannelInfoByID Lib "TSRemote" (ID As Long, ChannelInfo As tsrChannelInfoGet, PlayerInfo As Long, Records As Long) As Long
Private Declare Function tsrSwitchChannelID Lib "TSRemote" (ID As Long, Passwort As String) As Long
Private Declare Function tsrGetUserInfo Lib "TSRemote" (UserInfo As tsrUserInfo) As Long
Private Declare Function tsrSetPlayerFlags Lib "TSRemote" (PlayerFlags As Long) As Long
Private Declare Function tsrSendTextMessageToChannel Lib "TSRemote" (ID As Long, Text As String) As Long
Private Declare Function tsrKickPlayerFromChannel Lib "TSRemote" (ID As Long, Reason As String) As Long
Private Declare Function tsrKickPlayerFromServer Lib "TSRemote" (ID As Long, Reason As String) As Long
Private Declare Function tsrGetChannels Lib "TSRemote" (ByVal ChannelsInfo As Long, Records As Long) As Long
Private Declare Function tsrGetPlayers Lib "TSRemote" (ByVal PlayersInfo As Long, Records As Long) As Long
Private Declare Function tsrGetSpeakers Lib "TSRemote" (ByVal IDs As Long, Records As Long) As Long

'#######################
'#### Registry Functions
'#######################
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'######################
'#### Windows Functions
'######################
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'###########################
'#### Teamspeak Remote Const
'###########################

'#### codecs
Private Const cCodecCelp51 = 0
Private Const cCodecCelp63 = 1
Private Const cCodecGSM148 = 2
Private Const cCodecGSM164 = 3
Private Const cCodecWindowsCELP52 = 4
Private Const cCodecSpeex34 = 5
Private Const cCodecSpeex52 = 6
Private Const cCodecSpeex72 = 7
Private Const cCodecSpeex93 = 8
Private Const cCodecSpeex123 = 9
Private Const cCodecSpeex163 = 10
Private Const cCodecSpeex195 = 11
Private Const cCodecSpeex259 = 12

'#### codec masks
Private Const cmCelp51 = 2 ^ cCodecCelp51
Private Const cmCelp63 = 2 ^ cCodecCelp63
Private Const cmGSM148 = 2 ^ cCodecGSM148
Private Const cmGSM164 = 2 ^ cCodecGSM164
Private Const cmWindowsCELP52 = 2 ^ cCodecWindowsCELP52
Private Const cmSpeex34 = 2 ^ cCodecSpeex34
Private Const cmSpeex52 = 2 ^ cCodecSpeex52
Private Const cmSpeex72 = 2 ^ cCodecSpeex72
Private Const cmSpeex93 = 2 ^ cCodecSpeex93
Private Const cmSpeex123 = 2 ^ cCodecSpeex123
Private Const cmSpeex163 = 2 ^ cCodecSpeex163
Private Const cmSpeex195 = 2 ^ cCodecSpeex195
Private Const cmSpeex259 = 2 ^ cCodecSpeex259

'#### ServerType Flags
Private Const stClan = 1
Private Const stPublic = 2
Private Const stFreeware = 4
Private Const stCommercial = 8

'#### PlayerChannelPrivileges
Private Const pcpAdmin = 1
Private Const pcpOperator = 2
Private Const pcpAutoOperator = 4
Private Const pcpVoiced = 8
Private Const pcpAutoVoice = 16

'#### PlayerPrivileges
Private Const ppSuperServerAdmin = 1
Private Const ppServerAdmin = 2
Private Const ppCanRegister = 4
Private Const ppRegistered = 8
Private Const ppUnregistered = 16

'#### player flags
Private Const pfChannelCommander = 1
Private Const pfWantVoice = 2
Private Const pfNoWhisper = 4
Private Const pfAway = 8
Private Const pfInputMuted = 16
Private Const pfOutputMuted = 32
Private Const pfRecording = 64

'#### channel flags
Private Const cfRegistered = 1
Private Const cfUnregistered = 2
Private Const cfModerated = 4
Private Const cfPassword = 8
Private Const cfHierarchical = 16
Private Const cfDefault = 32

'#### Teamspeak Client
Private Const tsClientText = "TeamSpeak 2"
Private Const tsStatusClass = "TRICHEDITWITHLINKS"

Public Const tscJoin = "joined"
Public Const tscQuit = "quit"
Public Const tscSwitch = "switched"

'###################
'#### Registry Const
'###################
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_QUERY_VALUE = &H1
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0

'##################
'#### Windows Const
'##################
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE

'#################
'#### global lists
'#################
Public TSChannelList() As tsrChannelInfo
Public TSUserList() As tsrPlayerInfo
Public TSSpeakerIDs() As Long

'################
'#### global vars
'################
Public TSClientHandle As Long
Public TSClientStatusHandle As Long

'#####################################
'#### Teamspeak Remote Modul Functions
'#####################################

'#### TSGetLastError - Shows the last error
Public Sub TSGetLastError()
 Dim ErrorMessage As String * 1024
 
 tsrGetLastError ByVal ErrorMessage, ByVal 1024
 MsgBox ErrorMessage
End Sub

'#### TSConnect - Connect to a TS-Server
Public Sub TSConnect(ByVal server As String, nick As String, userID As String, pwd As String, _
                     Optional sChannel As String = "", Optional sChannelPW As String = "")
 Dim lRet As Long
 
 lRet = tsrConnect(ByVal "teamspeak://" & server & "?nickname=" & nick & "?loginname=" & userID & _
                   "?password=" & pwd & "?channel=" & sChannel & "?channelpassword=" & sChannelPW)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSDisconnect - Disconnect from a TS-Server
Public Sub TSDisconnect()
 Dim lRet As Long
 
 lRet = tsrDisconnect
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSQuit - Quit TS-Client
Public Sub TSQuit()
 Dim lRet As Long

 lRet = tsrQuit
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSGetClientVersion - Get Version of TS-Client
Public Function TSGetClientVersion() As String
 Dim tVer As tsrVersion
 
 tVer = tsrGetVersion
 TSGetClientVersion = tVer.Major & "." & tVer.Minor & "." & tVer.Release & "." & tVer.Build
End Function

'#### TSSwitchChannelByID - Switch to the channel with the ID
'####   - ChannelID       = ID of the channel
'####   - ChannelPasswort = Password for the channel
Public Sub TSSwitchChannelByID(ByVal ChannelID As Long, Optional ChannelPasswort As String = "")
 Dim lRet As Long
 
 lRet = tsrSwitchChannelID(ByVal ChannelID, ByVal ChannelPasswort)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSMuteSound - Mute sound
'####   - bMute = Yes/No
Public Sub TSMuteSound(ByVal bMute As Boolean)
 Dim lRet As Long
 Dim UInfo As tsrUserInfo
 Dim lFlags As Long
 
 lRet = tsrGetUserInfo(UInfo)
 If lRet <> 0 Then TSGetLastError

 lFlags = FourByteToLong(UInfo.Player.PlayerFlags(0), UInfo.Player.PlayerFlags(1), UInfo.Player.PlayerFlags(2), UInfo.Player.PlayerFlags(3))
 
 If bMute = True Then lFlags = lFlags Or (pfInputMuted Or pfOutputMuted) _
  Else lFlags = lFlags And Not (pfInputMuted Or pfOutputMuted)

 lRet = tsrSetPlayerFlags(ByVal lFlags)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSSetAway - Sets your status to away
'####   - bAway = Yes/No
Public Sub TSSetAway(ByVal bAway As Boolean)
 Dim lRet As Long
 Dim UInfo As tsrUserInfo
 Dim lFlags As Long
 
 lRet = tsrGetUserInfo(UInfo)
 If lRet <> 0 Then TSGetLastError

 lFlags = FourByteToLong(UInfo.Player.PlayerFlags(0), UInfo.Player.PlayerFlags(1), UInfo.Player.PlayerFlags(2), UInfo.Player.PlayerFlags(3))
 
 If bAway = True Then lFlags = lFlags Or pfAway _
  Else lFlags = lFlags And Not pfAway

 lRet = tsrSetPlayerFlags(ByVal lFlags)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSSendMessageToChannel - Sends a Message to the channel
'####   - ChannelID = ID of the channel
'####   - sText     = Text to send to the channel
Public Sub TSSendMessageToChannel(ByVal ChannelID As Long, sText As String)
 Dim lRet As Long
 
 lRet = tsrSendTextMessageToChannel(ByVal ChannelID, ByVal sText)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSKickUserFromChannel - Kicks a user from a channel
'####   - UserID  = ID of the user
'####   - sReason = Text for the user to see, why he/she was kicked
Public Sub TSKickUserFromChannel(ByVal userID As Long, sReason As String)
 Dim lRet As Long
 
 lRet = tsrKickPlayerFromChannel(ByVal userID, ByVal sReason)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSKickUserFromServer - Kicks a user from the server
'####   - UserID  = ID of the user
'####   - sReason = Text for the user to see, why he/she was kicked
Public Sub TSKickUserFromServer(ByVal userID As Long, sReason As String)
 Dim lRet As Long
 
 lRet = tsrKickPlayerFromServer(ByVal userID, ByVal sReason)
 If lRet <> 0 Then TSGetLastError
End Sub

'#### TSGetServerInfo - Get infomation for the connectet server
Public Function TSGetServerInfo() As tsrServerInfo
 Dim tSInfo As tsrServerInfoGet
 Dim lWert As Long
 Dim sInfo As String

 tSInfo = tsrGetServerInfo

 TSGetServerInfo.ServerName = StrConv(tSInfo.ServerName, vbUnicode)
 TSGetServerInfo.WelcomeMessage = StrConv(tSInfo.WelcomeMessage, vbUnicode)
 
 TSGetServerInfo.ServerVersion = FourByteToLong(tSInfo.ServerVMajor(0), tSInfo.ServerVMajor(1), tSInfo.ServerVMajor(2), tSInfo.ServerVMajor(3)) & "." & _
                                 FourByteToLong(tSInfo.ServerVMinor(0), tSInfo.ServerVMinor(1), tSInfo.ServerVMinor(2), tSInfo.ServerVMinor(3)) & "." & _
                                 FourByteToLong(tSInfo.ServerVRelease(0), tSInfo.ServerVRelease(1), tSInfo.ServerVRelease(2), tSInfo.ServerVRelease(3)) & "." & _
                                 FourByteToLong(tSInfo.ServerVBuild(0), tSInfo.ServerVBuild(1), tSInfo.ServerVBuild(2), tSInfo.ServerVBuild(3))
 
 TSGetServerInfo.ServerPlatform = StrConv(tSInfo.ServerPlatform, vbUnicode)
 TSGetServerInfo.ServerIp = StrConv(tSInfo.ServerIp, vbUnicode)
 TSGetServerInfo.ServerHost = StrConv(tSInfo.ServerHost, vbUnicode)
 
 sInfo = ""
 lWert = FourByteToLong(tSInfo.ServerType(0), tSInfo.ServerType(1), tSInfo.ServerType(2), tSInfo.ServerType(3))
 If ((lWert And stFreeware) <> 0) Then sInfo = "Freeware"
 If ((lWert And stCommercial) <> 0) Then sInfo = "Commercial"
 If ((lWert And stClan) <> 0) Then sInfo = sInfo & " Clan server"
 If ((lWert And stPublic) <> 0) Then sInfo = sInfo & " Public Server"
 TSGetServerInfo.ServerType = sInfo
 
 TSGetServerInfo.ServerMaxUsers = FourByteToLong(tSInfo.ServerMaxUsers(0), tSInfo.ServerMaxUsers(1), tSInfo.ServerMaxUsers(2), tSInfo.ServerMaxUsers(3))
 
 sInfo = ""
 lWert = FourByteToLong(tSInfo.SupportedCodecs(0), tSInfo.SupportedCodecs(1), tSInfo.SupportedCodecs(2), tSInfo.SupportedCodecs(3))
 If (cmCelp51 And lWert <> 0) Then sInfo = "Celp 5.1, "
 If (cmCelp63 And lWert <> 0) Then sInfo = sInfo & "Celp 6.3, "
 If (cmGSM148 And lWert <> 0) Then sInfo = sInfo & "GSM 14.8, "
 If (cmGSM164 And lWert <> 0) Then sInfo = sInfo & "GSM 16.4, "
 If (cmWindowsCELP52 And lWert <> 0) Then sInfo = sInfo & "WCELP 5.2, "
 If (cmSpeex34 And lWert <> 0) Then sInfo = sInfo & "Speex 3.4, "
 If (cmSpeex52 And lWert <> 0) Then sInfo = sInfo & "Speex 5.2, "
 If (cmSpeex72 And lWert <> 0) Then sInfo = sInfo & "Speex 7.2, "
 If (cmSpeex93 And lWert <> 0) Then sInfo = sInfo & "Speex 9.3, "
 If (cmSpeex123 And lWert <> 0) Then sInfo = sInfo & "Speex 12.3, "
 If (cmSpeex163 And lWert <> 0) Then sInfo = sInfo & "Speex 16.3, "
 If (cmSpeex195 And lWert <> 0) Then sInfo = sInfo & "Speex 19.5, "
 If (cmSpeex259 And lWert <> 0) Then sInfo = sInfo & "Speex 25.9, "
 If sInfo <> "" Then sInfo = Mid(sInfo, 1, Len(sInfo) - 2)
 TSGetServerInfo.SupportedCodecs = sInfo
 
 TSGetServerInfo.ChannelCount = FourByteToLong(tSInfo.ChannelCount(0), tSInfo.ChannelCount(1), tSInfo.ChannelCount(2), tSInfo.ChannelCount(3))
 TSGetServerInfo.PlayerCount = FourByteToLong(tSInfo.PlayerCount(0), tSInfo.PlayerCount(1), tSInfo.PlayerCount(2), tSInfo.PlayerCount(3))
End Function

'#### TSGetPlayerInfoByID - Get infomation for the user with ID
'####   - PlayerID = ID of the user
Public Function TSGetPlayerInfoByID(ByVal PlayerID As Long) As tsrPlayerInfo
 Dim lRet As Long
 Dim PlayerInfo As tsrPlayerInfoGet
 Dim lWert As Long
 Dim sInfo As String
 
 lRet = tsrGetPlayerInfoByID(ByVal PlayerID, PlayerInfo)
 If lRet <> 0 Then TSGetLastError
 
 TSGetPlayerInfoByID.PlayerID = FourByteToLong(PlayerInfo.PlayerID(0), PlayerInfo.PlayerID(1), PlayerInfo.PlayerID(2), PlayerInfo.PlayerID(3))
 TSGetPlayerInfoByID.ChannelID = FourByteToLong(PlayerInfo.ChannelID(0), PlayerInfo.ChannelID(1), PlayerInfo.ChannelID(2), PlayerInfo.ChannelID(3))
 TSGetPlayerInfoByID.NickName = StrConv(PlayerInfo.NickName, vbUnicode)
 
 sInfo = ""
 lWert = FourByteToLong(PlayerInfo.PlayerChannelPrivileges(0), PlayerInfo.PlayerChannelPrivileges(1), PlayerInfo.PlayerChannelPrivileges(2), PlayerInfo.PlayerChannelPrivileges(3))
 If ((lWert And pcpAdmin) <> 0) Then sInfo = "Admin,"
 If ((lWert And pcpOperator) <> 0) Then sInfo = sInfo & "Operator,"
 If ((lWert And pcpAutoOperator) <> 0) Then sInfo = sInfo & "Auto-Operator,"
 If ((lWert And pcpVoiced) <> 0) Then sInfo = sInfo & "Voice,"
 If ((lWert And pcpAutoVoice) <> 0) Then sInfo = sInfo & "Auto-Voice,"
 If sInfo <> "" Then sInfo = Mid(sInfo, 1, Len(sInfo) - 1) Else sInfo = "none"
 TSGetPlayerInfoByID.PlayerChannelPrivileges = sInfo

 sInfo = ""
 lWert = FourByteToLong(PlayerInfo.PlayerPrivileges(0), PlayerInfo.PlayerPrivileges(1), PlayerInfo.PlayerPrivileges(2), PlayerInfo.PlayerPrivileges(3))
 If ((lWert And ppSuperServerAdmin) <> 0) Then sInfo = "SuperServerAdmin,"
 If ((lWert And ppServerAdmin) <> 0) Then sInfo = sInfo & "ServerAdmin,"
 If ((lWert And ppCanRegister) <> 0) Then sInfo = sInfo & "CanRegister,"
 If ((lWert And ppRegistered) <> 0) Then sInfo = sInfo & "Registered" _
  Else sInfo = sInfo & "Unregistered"
 TSGetPlayerInfoByID.PlayerPrivileges = sInfo

 sInfo = ""
 lWert = FourByteToLong(PlayerInfo.PlayerFlags(0), PlayerInfo.PlayerFlags(1), PlayerInfo.PlayerFlags(2), PlayerInfo.PlayerFlags(3))
 If ((lWert And pfChannelCommander) <> 0) Then sInfo = "ChannelCommander,"
 If ((lWert And pfWantVoice) <> 0) Then sInfo = sInfo & "WantVoice,"
 If ((lWert And pfNoWhisper) <> 0) Then sInfo = sInfo & "NoWhisper,"
 If ((lWert And pfAway) <> 0) Then sInfo = sInfo & "Away,"
 If ((lWert And pfInputMuted) <> 0) Then sInfo = sInfo & "InputMuted,"
 If ((lWert And pfOutputMuted) <> 0) Then sInfo = sInfo & "OutputMuted,"
 If ((lWert And pfRecording) <> 0) Then sInfo = sInfo & "Recording,"
 If sInfo <> "" Then sInfo = Mid(sInfo, 1, Len(sInfo) - 1) Else sInfo = "none"
 TSGetPlayerInfoByID.PlayerFlags = sInfo
End Function

'#### TSGetChannelList - Save the Channellist of the server in global array TSChannelList()
'####   - Function returns TRUE if channellist is ok
Public Function TSGetChannelList() As Boolean
 Dim lRet As Long
 Dim ChannelsInfo() As tsrChannelInfoGet
 Dim lRecords As Long
 Dim i As Integer
 Dim sInfo As String
 Dim lWert As Long

 TSGetChannelList = False
 
 ReDim ChannelsInfo(1023)

 lRecords = 1024
 lRet = tsrGetChannels(VarPtr(ChannelsInfo(0)), lRecords)
 If lRet <> 0 Then
  TSGetLastError
  Exit Function
 End If
 
 ReDim Preserve ChannelsInfo(lRecords - 1)
 ReDim TSChannelList(lRecords - 1)
 
 For i = 0 To lRecords - 1
  TSChannelList(i).ChannelName = StrConv(ChannelsInfo(i).ChannelName, vbUnicode)
  TSChannelList(i).ChannelID = FourByteToLong(ChannelsInfo(i).ChannelID(0), ChannelsInfo(i).ChannelID(1), ChannelsInfo(i).ChannelID(2), ChannelsInfo(i).ChannelID(3))
  TSChannelList(i).ChannelParentID = FourByteToLong(ChannelsInfo(i).ChannelParentID(0), ChannelsInfo(i).ChannelParentID(1), ChannelsInfo(i).ChannelParentID(2), ChannelsInfo(i).ChannelParentID(3))
  TSChannelList(i).PlayerCountInChannel = FourByteToLong(ChannelsInfo(i).PlayerCountInChannel(0), ChannelsInfo(i).PlayerCountInChannel(1), ChannelsInfo(i).PlayerCountInChannel(2), ChannelsInfo(i).PlayerCountInChannel(3))
 
  sInfo = ""
  lWert = FourByteToLong(ChannelsInfo(i).ChannelFlags(0), ChannelsInfo(i).ChannelFlags(1), ChannelsInfo(i).ChannelFlags(2), ChannelsInfo(i).ChannelFlags(3))
  If ((lWert And cfRegistered) <> 0) Then sInfo = sInfo & "Registered,"
  If ((lWert And cfUnregistered) <> 0) Then sInfo = sInfo & "Unregistered,"
  If ((lWert And cfModerated) <> 0) Then sInfo = sInfo & "Moderated,"
  If ((lWert And cfPassword) <> 0) Then sInfo = sInfo & "Password,"
  If ((lWert And cfHierarchical) <> 0) Then sInfo = sInfo & "Hierarchical,"
  If ((lWert And cfDefault) <> 0) Then sInfo = sInfo & "Default,"
  sInfo = Mid(sInfo, 1, Len(sInfo) - 1)
  TSChannelList(i).ChannelFlags = sInfo
 
  lWert = FourByteToLong(ChannelsInfo(i).Codec(0), ChannelsInfo(i).Codec(1), ChannelsInfo(i).Codec(2), ChannelsInfo(i).Codec(3))
  Select Case lWert
   Case cCodecCelp51: sInfo = "CELP 5.1 Kbit"
   Case cCodecCelp63: sInfo = "CELP 6.3 Kbit"
   Case cCodecGSM148: sInfo = "GSM 14.8 Kbit"
   Case cCodecGSM164: sInfo = "GSM 16.4 Kbit"
   Case cCodecWindowsCELP52: sInfo = "CELP Windows 5.2 Kbit"
   Case cCodecSpeex34: sInfo = "Speex 3.4 Kbit"
   Case cCodecSpeex52: sInfo = "Speex 5.2 Kbit"
   Case cCodecSpeex72: sInfo = "Speex 7.2 Kbit"
   Case cCodecSpeex93: sInfo = "Speex 9.3 Kbit"
   Case cCodecSpeex123: sInfo = "Speex 12.3 Kbit"
   Case cCodecSpeex163: sInfo = "Speex 16.3 Kbit"
   Case cCodecSpeex195: sInfo = "Speex 19.5 Kbit"
   Case cCodecSpeex259: sInfo = "Speex 25.9 Kbit"
  End Select
  TSChannelList(i).Codec = sInfo
 Next i
 
 TSGetChannelList = True
End Function

'#### TSGetUserList - Save the Userlist of the server in global array TSUserList()
'####   - Function returns TRUE if userlist is ok
Public Function TSGetUserList() As Boolean
 Dim lRet As Long
 Dim PlayersInfo() As tsrPlayerInfoGet
 Dim lRecords As Long
 Dim i As Integer
 Dim sInfo As String
 Dim lWert As Long

 TSGetUserList = False

 ReDim PlayersInfo(1023)

 lRecords = 1024
 lRet = tsrGetPlayers(VarPtr(PlayersInfo(0)), lRecords)
 If lRet <> 0 Then
  TSGetLastError
  Exit Function
 End If
 
 ReDim Preserve PlayersInfo(lRecords - 1)
 ReDim TSUserList(lRecords - 1)
 
 For i = 0 To lRecords - 1
  TSUserList(i).NickName = StrConv(PlayersInfo(i).NickName, vbUnicode)
  TSUserList(i).PlayerID = FourByteToLong(PlayersInfo(i).PlayerID(0), PlayersInfo(i).PlayerID(1), PlayersInfo(i).PlayerID(2), PlayersInfo(i).PlayerID(3))
  TSUserList(i).ChannelID = FourByteToLong(PlayersInfo(i).ChannelID(0), PlayersInfo(i).ChannelID(1), PlayersInfo(i).ChannelID(2), PlayersInfo(i).ChannelID(3))
 
  sInfo = ""
  lWert = FourByteToLong(PlayersInfo(i).PlayerChannelPrivileges(0), PlayersInfo(i).PlayerChannelPrivileges(1), PlayersInfo(i).PlayerChannelPrivileges(2), PlayersInfo(i).PlayerChannelPrivileges(3))
  If ((lWert And pcpAdmin) <> 0) Then sInfo = "Admin,"
  If ((lWert And pcpOperator) <> 0) Then sInfo = sInfo & "Operator,"
  If ((lWert And pcpAutoOperator) <> 0) Then sInfo = sInfo & "Auto-Operator,"
  If ((lWert And pcpVoiced) <> 0) Then sInfo = sInfo & "Voice,"
  If ((lWert And pcpAutoVoice) <> 0) Then sInfo = sInfo & "Auto-Voice,"
  If sInfo <> "" Then sInfo = Mid(sInfo, 1, Len(sInfo) - 1) Else sInfo = "none"
  TSUserList(i).PlayerChannelPrivileges = sInfo

  sInfo = ""
  lWert = FourByteToLong(PlayersInfo(i).PlayerPrivileges(0), PlayersInfo(i).PlayerPrivileges(1), PlayersInfo(i).PlayerPrivileges(2), PlayersInfo(i).PlayerPrivileges(3))
  If ((lWert And ppSuperServerAdmin) <> 0) Then sInfo = "SuperServerAdmin,"
  If ((lWert And ppServerAdmin) <> 0) Then sInfo = sInfo & "ServerAdmin,"
  If ((lWert And ppCanRegister) <> 0) Then sInfo = sInfo & "CanRegister,"
  If ((lWert And ppRegistered) <> 0) Then sInfo = sInfo & "Registered" _
   Else sInfo = sInfo & "Unregistered"
  TSUserList(i).PlayerPrivileges = sInfo

  sInfo = ""
  lWert = FourByteToLong(PlayersInfo(i).PlayerFlags(0), PlayersInfo(i).PlayerFlags(1), PlayersInfo(i).PlayerFlags(2), PlayersInfo(i).PlayerFlags(3))
  If ((lWert And pfChannelCommander) <> 0) Then sInfo = "ChannelCommander,"
  If ((lWert And pfWantVoice) <> 0) Then sInfo = sInfo & "WantVoice,"
  If ((lWert And pfNoWhisper) <> 0) Then sInfo = sInfo & "NoWhisper,"
  If ((lWert And pfAway) <> 0) Then sInfo = sInfo & "Away,"
  If ((lWert And pfInputMuted) <> 0) Then sInfo = sInfo & "InputMuted,"
  If ((lWert And pfOutputMuted) <> 0) Then sInfo = sInfo & "OutputMuted,"
  If ((lWert And pfRecording) <> 0) Then sInfo = sInfo & "Recording,"
  If sInfo <> "" Then sInfo = Mid(sInfo, 1, Len(sInfo) - 1) Else sInfo = "none"
  TSUserList(i).PlayerFlags = sInfo
 Next i
 
 TSGetUserList = True
End Function

'#### TSGetSpeakers - Save the IDs of the speaking users in global array TSSpeakerIDs()
'####   - Function returns the actual Status of Client und Connection
Public Function TSGetSpeakers() As tsStatus
 Dim lRet As Long
 Dim lRecords As Long
 Dim i As Integer
 Dim sInfo As String
 Dim lWert As Long
 
 TSGetSpeakers = TSUNKNOWN
 
 ReDim TSSpeakerIDs(1023)
 
 lRecords = 1024
 lRet = tsrGetSpeakers(VarPtr(TSSpeakerIDs(0)), lRecords)
 Select Case lRet
  Case -1000
   TSGetSpeakers = ClientNotRunning
   Exit Function
  Case -1
   TSGetSpeakers = NotConnected
   Exit Function
  Case 0
   'All OK
  Case Else
   TSGetLastError
   Exit Function
 End Select

 TSGetSpeakers = Connected
 If lRecords = 0 Then Exit Function
 ReDim Preserve TSSpeakerIDs(lRecords - 1)
End Function

'#############################################
'#### Teamspeak Remote Modul Helping Functions
'#############################################

'#### Function to get a Long from 4 Bytes
Private Function FourByteToLong(LoByte1 As Byte, HiByte1 As Byte, LoByte2 As Byte, HiByte2 As Byte) As Long
 Dim LoWord As Integer, HiWord As Integer
 
 If HiByte1 And &H80 Then LoWord1 = ((HiByte1 * &H100&) Or LoByte1) Or &HFFFF0000 _
  Else LoWord = (HiByte1 * &H100) Or LoByte1
 If HiByte2 And &H80 Then LoWord2 = ((HiByte2 * &H100&) Or LoByte2) Or &HFFFF0000 _
  Else HiWord = (HiByte2 * &H100) Or LoByte2

 FourByteToLong = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

'#### Function to get 4 Bytes from Long
'Private Function LongToFourBytes(ByVal LongWert As Long) As Variant
' Dim HiWord As Integer, LoWord As Integer
' Dim LoByte1 As Byte, HiByte1 As Byte, LoByte2 As Byte, HiByte2 As Byte
' Dim vBytes As Variant
'
' HiWord = (LongWert And &HFFFF0000) \ &H10000
' If LongWert And &H8000& Then LoWord = LongWert Or &HFFFF0000 _
'  Else LoWord = LongWert And &HFFFF&
'
' HiByte1 = (LoWord And &HFF00&) \ &H100
' LoByte1 = LoWord And &HFF
' HiByte2 = (HiWord And &HFF00&) \ &H100
' LoByte2 = HiWord And &HFF
'
' ReDim vBytes(3)
' vBytes(0) = LoByte1
' vBytes(1) = HiByte1
' vBytes(2) = LoByte2
' vBytes(3) = HiByte2
' LongToFourBytes = vBytes
'End Function

'################################################
'#### Teamspeak Remote Modul Additional Functions
'################################################

'#### TSGetClientPath - Reads the path of the TS-Client from registry
Public Function TSGetClientPath() As String
 Dim sKeyName As String
 Dim sBuffer As String
 Dim bLen As Long
 Dim lType As Long
 Dim lHandle As Long
 Dim lResult As Long
 Dim nCount As Long
 Dim tmp(0 To 254) As Byte

 lResult = RegOpenKeyEx(HKEY_CLASSES_ROOT, "teamspeak\DefaultIcon", 0&, KEY_QUERY_VALUE, lHandle)
 If lResult = ERROR_SUCCESS Then

  sKeyName = Space(255)
  lResult = RegEnumValue(lHandle, nCount, sKeyName, Len(sKeyName), 0&, 0&, tmp(0), 256)
  If lResult <> ERROR_SUCCESS Then GoTo CloseHandle

  sKeyName = Left$(sKeyName, InStr(sKeyName, vbNullChar) - 1)

  sBuffer = Space$(255)
  bLen = Len(sBuffer)
  lType = REG_SZ
  lResult = RegQueryValueEx(lHandle, sKeyName, 0&, lType, ByVal sBuffer, bLen)
  
  If lResult = ERROR_SUCCESS Then
   sBuffer = Left$(sBuffer, bLen)
   While Right$(sBuffer, 1) = Chr$(0)
    sBuffer = Left$(sBuffer, Len(sBuffer) - 1)
   Wend
  Else
   sBuffer = ""
  End If

CloseHandle:
  RegCloseKey lHandle
 End If
 
 TSGetClientPath = sBuffer
End Function

'#### TSGetClientWindowAndStatus - Get the Handle of TS-Client and Status-Window
Public Sub TSGetClientWindowAndStatusHandle()
 TSClientHandle = FindWindow(vbNullString, tsClientText)
 
 If TSClientHandle <> 0 Then Call EnumChildWindows(TSClientHandle, AddressOf WndEnumChildStatus, 0&) _
  Else TSClientStatusHandle = 0
End Sub

Private Function WndEnumChildStatus(ByVal hwnd As Long, lParam As Long) As Long
 Dim sClassName As String * 50
 Dim sWindowText As String

 Call GetClassName(hwnd, sClassName, 50)
 If UCase(Left(sClassName, 18)) = tsStatusClass Then
  sWindowText = GetText(hwnd)
  If (Trim(sWindowText) <> "") And _
     (Left$(sWindowText, 7) <> "Server:") And _
     (Left$(sWindowText, 8) <> "Channel:") And _
     (Left$(sWindowText, 7) <> "Player:") Then TSClientStatusHandle = hwnd
 End If
 
 WndEnumChildStatus = 1
End Function

Private Function GetText(ByVal lHandle As Long) As String
 Dim Textlen As Long
 Dim Text As String

 Textlen = SendMessage(lHandle, WM_GETTEXTLENGTH, 0, 0)
 If Textlen = 0 Then
  GetText = ""
  Exit Function
 End If

 Textlen = Textlen + 1
 Text = Space(Textlen)
 Textlen = SendMessage(lHandle, WM_GETTEXT, Textlen, ByVal Text)
 GetText = Left(Text, Textlen)
End Function

'#### TSGetStatusTextLen - Get the len of the Statustext in TS-Client
Public Function TSGetStatusTextLen() As Long
 TSGetStatusTextLen = SendMessage(TSClientStatusHandle, WM_GETTEXTLENGTH, 0, 0)
End Function

'#### TSGetStatusText - Get the Text of the Statustext in TS-Client
Public Function TSGetStatusText() As String
 TSGetStatusText = GetText(TSClientStatusHandle)
End Function

Public Function TS2Installed() As Boolean
    Dim TS2Path As String
    
    TS2Path = RegReadString(HKEY_CLASSES_ROOT, "teamspeak\Shell\Open\command", "", "")
    TS2Installed = (TS2Path <> "")
End Function

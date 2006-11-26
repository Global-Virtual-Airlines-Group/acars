Attribute VB_Name = "modRegistry"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
    (ByVal hKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" _
    Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
    Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey _
    As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal _
    lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" _
    Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
    Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName _
    As String, ByVal lpReserved As Long, lpType As Long, lpData _
    As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
    Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal Reserved As Long, ByVal _
    dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Global Const KEY_QUERY_VALUE = &H1
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003
Global Const HKCU = HKEY_CURRENT_USER
Global Const HKLM = HKEY_LOCAL_MACHINE

Public Const REG_SZ = 1
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const ERROR_SUCCESS = 0&

Public Sub RegDelValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegDeleteValue(hCurKey, strValue)
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub RegWriteKey(hKey As Long, strPath As String)
    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Sub DelKey(ByVal hKey As Long, ByVal strPath As String)
    Dim lRegResult As Long
    lRegResult = RegDeleteKey(hKey, strPath)
End Sub

Public Function RegReadString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String

    Dim hCurKey As Long
    Dim lResult As Long
    Dim lValueType As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim lRegResult As Long

    'Set up default value
    If Not IsEmpty(Default) Then
        RegReadString = Default
    Else
        RegReadString = ""
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_SZ Then
            strBuffer = String(lDataBufferSize, " ")
            lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)

            intZeroPos = InStr(strBuffer, Chr$(0))
            If intZeroPos > 0 Then
                RegReadString = Left$(strBuffer, intZeroPos - 1)
            Else
                RegReadString = strBuffer
            End If
        End If
    End If

    lRegResult = RegCloseKey(hCurKey)
End Function

Public Sub RegWriteString(hKey As Long, strPath As String, strValue As String, strData As String)

    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function RegReadBoolean(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Boolean = False) As Boolean
    Dim defaultLong As Long

    If Default Then
        defaultLong = 1
    Else
        defaultLong = 0
    End If

    RegReadBoolean = (RegReadDWORD(hKey, strPath, strValue, defaultLong) = 1)
End Function

Public Function RegReadDWORD(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long

    Dim lRegResult As Long
    Dim lValueType As Long
    Dim lBuffer As Long
    Dim lDataBufferSize As Long
    Dim hCurKey As Long
    
    'Set critical error handler
    On Error GoTo ErrorHandler

    'Set up default value
    If Not IsEmpty(Default) Then
        RegReadDWORD = Default
    Else
        RegReadDWORD = 0
    End If

    lRegResult = RegOpenKey(hKey, strPath, hCurKey)
    lDataBufferSize = 4 '4 bytes = 32 bits = long
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)

    If lRegResult = ERROR_SUCCESS Then
        If lValueType = REG_DWORD Then RegReadDWORD = lBuffer
    End If

    lRegResult = RegCloseKey(hCurKey)
    
ExitSub:
    Exit Function
    
ErrorHandler:
    RegReadDWORD = Default
    Resume ExitSub
    
End Function

Public Sub RegWriteBoolean(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Boolean)
    Dim keyValue As Long
    
    If lData Then
        keyValue = 1
    Else
        keyValue = 0
    End If
        
    Call RegWriteDWORD(hKey, strPath, strValue, keyValue)
End Sub

Public Sub RegWriteDWORD(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)

    Dim hCurKey As Long
    Dim lRegResult As Long

    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
    lRegResult = RegCloseKey(hCurKey)
End Sub

Function ReadINI(ByVal sSection As String, ByVal sKey As String, ByVal sDefault As String, ByVal sIniFile As String)
    Dim sBuffer As String, lRet As Long
    ' Fill String with 255 spaces
    sBuffer = String$(255, 0)
    
    ' Call DLL
    lRet = GetPrivateProfileString(sSection, sKey, "", sBuffer, Len(sBuffer), sIniFile)
    If lRet = 0 Then
        ' DLL failed, save default
        If sDefault <> "" Then WriteINI sSection, sKey, sDefault, sIniFile
        ReadINI = sDefault
        Exit Function
    End If
    
    'Load the value, and then find first whitespace
    ReadINI = Left(sBuffer, InStr(sBuffer, Chr(0)) - 1)
    If (InStr(ReadINI, ";") > 0) Then
        ReadINI = Left(ReadINI, InStr(ReadINI, ";") - 1)
    ElseIf (InStr(ReadINI, "//") > 0) Then
        ReadINI = Left(ReadINI, InStr(ReadINI, "//") - 1)
    End If
End Function

' Returns True if successful. If section does not exist it creates it.
Function WriteINI(sSection As String, sKey As String, sValue As String, sIniFile As String) As Boolean
    Dim lRet As Long
    
    ' Call DLL
    lRet = WritePrivateProfileString(sSection, sKey, sValue, sIniFile)
    WriteINI = (lRet)
End Function

Function GetINISections(ByVal sIniFile As String) As Variant
    Dim fNum As Integer
    Dim rawData As Variant
    Dim entry As Variant
    Dim results As Variant
    
    'Set critical error handler
    On Error GoTo FatalError
    
    'Get free file
    fNum = FreeFile()
    Open sIniFile For Input Lock Write As #fNum
    rawData = Split(Input(LOF(fNum), fNum), vbCrLf)
    Close #fNum
    
    'Get the sections
    For Each entry In rawData
        If ((Left(entry, 1) = "[") And (Right(entry, 1) = "]")) Then
            If IsEmpty(results) Then
                ReDim results(0)
            Else
                ReDim Preserve results(UBound(results) + 1)
            End If
            
            results(UBound(results)) = Mid(entry, 2, Len(entry) - 2)
        End If
    Next
    
ExitSub:
    GetINISections = results
    Exit Function
    
FatalError:
    Resume ExitSub
    
End Function

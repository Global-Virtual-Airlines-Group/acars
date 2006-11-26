Attribute VB_Name = "RadioUpdate"
Option Explicit

Private Const comRadioReset = 1
Private Const navRadioReset = 2

Private Function convertNAVCOM(ByVal freq As String) As Integer
    'returns 110.15 as 0x1015
    Dim freqParts As Variant
    Dim freqTop As Integer
    Dim freqBottom As Integer
    Dim retVal As Integer
    
    On Error GoTo FatalError
    
    'Check for empty string
    If (freq = "") Then
        convertNAVCOM = -1
        Exit Function
    End If
    
    'Split the frequency
    freqParts = Split(freq, ".")
    freqTop = CInt(freqParts(0)) - 100
    If (Len(freqParts(1)) = 1) Then freqParts(1) = freqParts(1) + "0"
    
    'Convert the frequency
    If (Len(freqParts(1)) = 1) Then
        freqParts(1) = freqParts(1) + "0"
    ElseIf (Len(freqParts(1)) > 2) Then
        freqParts(1) = Left(freqParts(1), 2)
    End If
    
    'Convert the fractional part to hex BCD
    freqBottom = CInt(freqParts(1))
    retVal = ((freqBottom \ 10) * 16) + (freqBottom Mod 10)
    
    'Convert the integer part to hex BCD
    retVal = retVal + ((freqTop Mod 10) * 256)
    convertNAVCOM = retVal + ((freqTop \ 10) * 4096)
    
ExitSub:
    Exit Function
    
FatalError:
    ShowMessage "Error processing Frequency " + freq, ACARSERRORCOLOR
    convertNAVCOM = -1
    Resume ExitSub
End Function

Private Function convertADF(ByVal freq As String) As Long
    'returns 1343.5 as 0x01050343
    Dim freqTop As Integer
    Dim retTop As Integer
    Dim retBottom As Integer
    
    On Error GoTo FatalError
    
    'Check for empty string
    If (freq = "") Then
        convertADF = -1
        Exit Function
    End If
    
    'Split the frequency
    freqTop = Fix(CDbl(freq))
    retTop = CInt((CDbl(freq) - freqTop) * 10) + ((freqTop \ 1000) * 256)

    'Convert the integer part to hex BCD
    retBottom = (((freqTop Mod 1000) \ 100) * 256) + (freqTop Mod 10)
    retBottom = retBottom + (((freqTop Mod 100) \ 10) * 16)
    convertADF = (retTop * 65536) + retBottom
    
ExitSub:
    Exit Function
    
FatalError:
    ShowMessage "Error processing Frequency " + freq, ACARSERRORCOLOR
    convertADF = -1
    Resume ExitSub
End Function

Private Function convertTX(ByVal freq As String) As Long
    Dim freqTop As Integer

    'Check for empty string
    If (freq = "") Then
        convertTX = -1
        Exit Function
    End If
    
    'Convert to BCD
    convertTX = Val("&H" + freq)
End Function

Public Sub SetNAV1(freqStr As String, hdg As Integer)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency/heading
    freq = convertNAVCOM(freqStr)
    If (freq = -1) Then Exit Sub
    Call FSUIPC_Write(&H350, 2, VarPtr(freq), lngResult)
    Call FSUIPC_Write(&HC4E, 2, VarPtr(hdg), lngResult)
    
    'Tell FS that the NAV freq changesd
    Call FSUIPC_Write(&H388, 1, VarPtr(navRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting NAV1", ACARSERRORCOLOR
    Else
        ShowMessage "NAV1 Radio set to " & freqStr, ACARSTEXTCOLOR
    End If
End Sub

Public Sub SetNAV2(freqStr As String)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency
    freq = convertNAVCOM(freqStr)
    If (freq = -1) Then Exit Sub
    Call FSUIPC_Write(&H352, 2, VarPtr(freq), lngResult)
                        
    'Tell FS that the NAV freq changesd
    Call FSUIPC_Write(&H388, 1, VarPtr(navRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting NAV2", ACARSERRORCOLOR
    Else
        ShowMessage "NAV2 Radio set to " & freqStr, ACARSTEXTCOLOR
    End If
End Sub

Public Sub SetCOM1(ByVal freqStr As String)
    Dim freq As Integer
    Dim lngResult As Long
    
    'Write the frequency
    freq = convertNAVCOM(freqStr)
    If (freq = -1) Then Exit Sub
    Call FSUIPC_Write(&H34E, 2, VarPtr(freq), lngResult)
    Call FSUIPC_Write(&H38A, 1, VarPtr(comRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting COM1", ACARSERRORCOLOR
    Else
        ShowMessage "COM1 Radio set to " & freqStr, ACARSTEXTCOLOR
    End If
End Sub

Public Sub SetCOM2(ByVal freqStr As String)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency
    freq = convertNAVCOM(freqStr)
    If (freq = -1) Then Exit Sub
    Call FSUIPC_Write(&H3118, 2, VarPtr(freq), lngResult)
    Call FSUIPC_Write(&H38A, 1, VarPtr(comRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting COM2", ACARSERRORCOLOR
    Else
        ShowMessage "COM2 Radio set to " & freqStr, ACARSTEXTCOLOR
    End If
End Sub

Public Sub SetADF1(ByVal freqStr As String)
    Dim freq As Long
    Dim fTop As Integer
    Dim fBottom As Integer
    Dim lngResult As Long
    
    'Write the frequency
    freq = convertADF(freqStr)
    If (freq = -1) Then Exit Sub
    
    'Split the frequency into its two offsets
    fTop = (freq \ 65536)
    fBottom = CInt(freq Mod 65536)
    
    'Write to FSUIPC
    Call FSUIPC_Write(&H34C, 2, VarPtr(fBottom), lngResult)
    Call FSUIPC_Write(&H356, 2, VarPtr(fTop), lngResult)
    Call FSUIPC_Write(&H389, 1, VarPtr(navRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting ADF1", ACARSERRORCOLOR
    Else
        ShowMessage "ADF1 Radio set to " & freqStr, ACARSTEXTCOLOR
    End If
End Sub

Public Sub SetTX(ByVal freqStr As String)
    Dim freq As Long
    Dim lngResult As Long
    
    'Write the frequency
    freq = convertTX(freqStr)
    If (freq = -1) Then Exit Sub

    'Write to FSUIPC
    Call FSUIPC_Write(&H354, 2, VarPtr(freq), lngResult)
    If Not FSUIPC_Process(lngResult) Then
        ShowMessage "Error setting Transponder Code", ACARSERRORCOLOR
    Else
        ShowMessage "Transponder set to " & freqStr, ACARSTEXTCOLOR
    End If
End Sub

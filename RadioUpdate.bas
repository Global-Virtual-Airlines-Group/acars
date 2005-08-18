Attribute VB_Name = "RadioUpdate"
Option Explicit

Private Const comRadioReset = 1
Private Const navRadioReset = 2

Private Function convertNAVCOM(ByVal freq As String) As Integer

    'returns 110.15 as 0x1015
    Dim freqTop As Integer
    Dim freqBottom As Integer
    Dim retVal As Integer
    
    'Split the frequency
    freqTop = Fix(CDbl(freq))
    freqBottom = CInt((CDbl(freq) - freqTop) * 100)
    
    'Convert the fractional part to hex BCD
    retVal = (Int(freqBottom / 10) * 16) + (freqBottom Mod 10)
    
    'Convert the integer part to hex BCD
    freqTop = freqTop - 100
    retVal = retVal + ((freqTop Mod 10) * 256)
    convertNAVCOM = retVal + (Int(freqTop / 10) * 4096)
End Function

Private Function convertADF(ByVal freq As String) As Long

    'returns 343.5 as 0x00050343
    Dim freqTop As Integer
    Dim freqBottom As Integer
    Dim retTop As Integer
    Dim retBottom As Integer
    
    'Split the frequency
    freqTop = Fix(CDbl(freq))
    freqBottom = CInt((CDbl(freq) - freqTop) * 100)

    'Convert the fractional part to hex BCD
    retTop = (Int(freqBottom / 10) * 16) + (freqBottom Mod 10)
    
    'Convert the integer part to hex BCD
    retBottom = (freqTop Mod 10) * 256
    retBottom = retBottom + (Int(freqTop / 100) * 4096)
    retBottom = retBottom + (Int(freqTop / 10) Mod 10) * 256
    convertADF = (retTop * 65536) + retBottom
End Function

Public Sub setNAV1(freqStr As String, hdg As Integer)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency/heading
    freq = convertNAVCOM(freqStr)
    Call FSUIPC_Write(&H350, 2, VarPtr(freq), lngResult)
    Call FSUIPC_Write(&HC4E, 2, VarPtr(hdg), lngResult)
    
    'Tell FS that the NAV freq changesd
    Call FSUIPC_Write(&H388, 1, VarPtr(navRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then ShowMessage "Error setting NAV1", ACARSERRORCOLOR
End Sub

Public Sub SetNAV2(freqStr As String)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency
    freq = convertNAVCOM(freqStr)
    Call FSUIPC_Write(&H352, 2, VarPtr(freq), lngResult)
                        
    'Tell FS that the NAV freq changesd
    Call FSUIPC_Write(&H388, 1, VarPtr(navRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then ShowMessage "Error setting NAV2", ACARSERRORCOLOR
End Sub

Public Sub SetCOM1(freqStr As String)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency
    freq = convertNAVCOM(freqStr)
    Call FSUIPC_Write(&H34E, 2, VarPtr(freq), lngResult)
    Call FSUIPC_Write(&H38A, 1, VarPtr(comRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then ShowMessage "Error setting COM1", ACARSERRORCOLOR
End Sub

Public Sub SetCOM2(freqStr As String)
    Dim freq As Integer
    Dim lngResult As Long

    'Write the frequency
    freq = convertNAVCOM(freqStr)
    Call FSUIPC_Write(&H3118, 2, VarPtr(freq), lngResult)
    Call FSUIPC_Write(&H38A, 1, VarPtr(comRadioReset), lngResult)
    If Not FSUIPC_Process(lngResult) Then ShowMessage "Error setting COM2", ACARSERRORCOLOR
End Sub

Public Sub setADF1(freqStr As String)
    Dim freq As Long
    Dim lngResult As Long
    
    'Write the frequency
    freq = convertADF(freqStr)
End Sub

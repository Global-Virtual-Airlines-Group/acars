VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SHA256"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' MODULE:       CSHA256
' FILENAME:     CSHA256.cls
' AUTHOR:       Phil Fresle
' CREATED:      10-Apr-2001
' COPYRIGHT:    Copyright 2001 Phil Fresle. All Rights Reserved.
'
' DESCRIPTION:
' This class is used to generate a SHA-256 'digest' or 'signature' of a string.
' The SHA-256 algorithm is one of the industry standard methods for generating
' digital signatures. It is generically known as a digest, digital signature,
' one-way encryption, hash or checksum algorithm. A common use for SHA-256 is
' for password encryption as it is one-way in nature, that does not mean that
' your passwords are not free from a dictionary attack. If you are using the
' routine for passwords, you can make it a little more secure by concatenating
' some known random characters to the password before you generate the signature
' and on subsequent tests, so even if a hacker knows you are using SHA-256 for
' your passwords, the random characters will make it harder to dictionary attack.
'
' *** CAUTION ***
' See the comment attached to the SHA256 method below regarding use on systems
' with different character sets.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
' Should you wish to commission some derivative work based on this code provided
' here, or any consultancy work, please do not hesitate to contact us.
'
' Web Site:  http://www.frez.co.uk
' E-mail:    sales@frez.co.uk
'
' MODIFICATION HISTORY:
' 10-Apr-2001   Phil Fresle     Initial Version
' 10-Feb-2006   Luke Kolin      Added salt support
'*******************************************************************************
Option Explicit

Private m_lOnBits(30)   As Long
Private m_l2Power(30)   As Long
Private K(63)           As Long

Private Const BITS_TO_A_BYTE  As Long = 8
Private Const BYTES_TO_A_WORD As Long = 4
Private Const BITS_TO_A_WORD  As Long = BYTES_TO_A_WORD * BITS_TO_A_BYTE

'*******************************************************************************
' Class_Initialize (SUB)
'*******************************************************************************
Private Sub Class_Initialize()
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - BITS SET FROM RIGHT
    m_lOnBits(0) = 1            ' 00000000000000000000000000000001
    m_lOnBits(1) = 3            ' 00000000000000000000000000000011
    m_lOnBits(2) = 7            ' 00000000000000000000000000000111
    m_lOnBits(3) = 15           ' 00000000000000000000000000001111
    m_lOnBits(4) = 31           ' 00000000000000000000000000011111
    m_lOnBits(5) = 63           ' 00000000000000000000000000111111
    m_lOnBits(6) = 127          ' 00000000000000000000000001111111
    m_lOnBits(7) = 255          ' 00000000000000000000000011111111
    m_lOnBits(8) = 511          ' 00000000000000000000000111111111
    m_lOnBits(9) = 1023         ' 00000000000000000000001111111111
    m_lOnBits(10) = 2047        ' 00000000000000000000011111111111
    m_lOnBits(11) = 4095        ' 00000000000000000000111111111111
    m_lOnBits(12) = 8191        ' 00000000000000000001111111111111
    m_lOnBits(13) = 16383       ' 00000000000000000011111111111111
    m_lOnBits(14) = 32767       ' 00000000000000000111111111111111
    m_lOnBits(15) = 65535       ' 00000000000000001111111111111111
    m_lOnBits(16) = 131071      ' 00000000000000011111111111111111
    m_lOnBits(17) = 262143      ' 00000000000000111111111111111111
    m_lOnBits(18) = 524287      ' 00000000000001111111111111111111
    m_lOnBits(19) = 1048575     ' 00000000000011111111111111111111
    m_lOnBits(20) = 2097151     ' 00000000000111111111111111111111
    m_lOnBits(21) = 4194303     ' 00000000001111111111111111111111
    m_lOnBits(22) = 8388607     ' 00000000011111111111111111111111
    m_lOnBits(23) = 16777215    ' 00000000111111111111111111111111
    m_lOnBits(24) = 33554431    ' 00000001111111111111111111111111
    m_lOnBits(25) = 67108863    ' 00000011111111111111111111111111
    m_lOnBits(26) = 134217727   ' 00000111111111111111111111111111
    m_lOnBits(27) = 268435455   ' 00001111111111111111111111111111
    m_lOnBits(28) = 536870911   ' 00011111111111111111111111111111
    m_lOnBits(29) = 1073741823  ' 00111111111111111111111111111111
    m_lOnBits(30) = 2147483647  ' 01111111111111111111111111111111
    
    ' Could have done this with a loop calculating each value, but simply
    ' assigning the values is quicker - POWERS OF 2
    m_l2Power(0) = 1            ' 00000000000000000000000000000001
    m_l2Power(1) = 2            ' 00000000000000000000000000000010
    m_l2Power(2) = 4            ' 00000000000000000000000000000100
    m_l2Power(3) = 8            ' 00000000000000000000000000001000
    m_l2Power(4) = 16           ' 00000000000000000000000000010000
    m_l2Power(5) = 32           ' 00000000000000000000000000100000
    m_l2Power(6) = 64           ' 00000000000000000000000001000000
    m_l2Power(7) = 128          ' 00000000000000000000000010000000
    m_l2Power(8) = 256          ' 00000000000000000000000100000000
    m_l2Power(9) = 512          ' 00000000000000000000001000000000
    m_l2Power(10) = 1024        ' 00000000000000000000010000000000
    m_l2Power(11) = 2048        ' 00000000000000000000100000000000
    m_l2Power(12) = 4096        ' 00000000000000000001000000000000
    m_l2Power(13) = 8192        ' 00000000000000000010000000000000
    m_l2Power(14) = 16384       ' 00000000000000000100000000000000
    m_l2Power(15) = 32768       ' 00000000000000001000000000000000
    m_l2Power(16) = 65536       ' 00000000000000010000000000000000
    m_l2Power(17) = 131072      ' 00000000000000100000000000000000
    m_l2Power(18) = 262144      ' 00000000000001000000000000000000
    m_l2Power(19) = 524288      ' 00000000000010000000000000000000
    m_l2Power(20) = 1048576     ' 00000000000100000000000000000000
    m_l2Power(21) = 2097152     ' 00000000001000000000000000000000
    m_l2Power(22) = 4194304     ' 00000000010000000000000000000000
    m_l2Power(23) = 8388608     ' 00000000100000000000000000000000
    m_l2Power(24) = 16777216    ' 00000001000000000000000000000000
    m_l2Power(25) = 33554432    ' 00000010000000000000000000000000
    m_l2Power(26) = 67108864    ' 00000100000000000000000000000000
    m_l2Power(27) = 134217728   ' 00001000000000000000000000000000
    m_l2Power(28) = 268435456   ' 00010000000000000000000000000000
    m_l2Power(29) = 536870912   ' 00100000000000000000000000000000
    m_l2Power(30) = 1073741824  ' 01000000000000000000000000000000
    
    ' Just put together the K array once
    K(0) = &H428A2F98
    K(1) = &H71374491
    K(2) = &HB5C0FBCF
    K(3) = &HE9B5DBA5
    K(4) = &H3956C25B
    K(5) = &H59F111F1
    K(6) = &H923F82A4
    K(7) = &HAB1C5ED5
    K(8) = &HD807AA98
    K(9) = &H12835B01
    K(10) = &H243185BE
    K(11) = &H550C7DC3
    K(12) = &H72BE5D74
    K(13) = &H80DEB1FE
    K(14) = &H9BDC06A7
    K(15) = &HC19BF174
    K(16) = &HE49B69C1
    K(17) = &HEFBE4786
    K(18) = &HFC19DC6
    K(19) = &H240CA1CC
    K(20) = &H2DE92C6F
    K(21) = &H4A7484AA
    K(22) = &H5CB0A9DC
    K(23) = &H76F988DA
    K(24) = &H983E5152
    K(25) = &HA831C66D
    K(26) = &HB00327C8
    K(27) = &HBF597FC7
    K(28) = &HC6E00BF3
    K(29) = &HD5A79147
    K(30) = &H6CA6351
    K(31) = &H14292967
    K(32) = &H27B70A85
    K(33) = &H2E1B2138
    K(34) = &H4D2C6DFC
    K(35) = &H53380D13
    K(36) = &H650A7354
    K(37) = &H766A0ABB
    K(38) = &H81C2C92E
    K(39) = &H92722C85
    K(40) = &HA2BFE8A1
    K(41) = &HA81A664B
    K(42) = &HC24B8B70
    K(43) = &HC76C51A3
    K(44) = &HD192E819
    K(45) = &HD6990624
    K(46) = &HF40E3585
    K(47) = &H106AA070
    K(48) = &H19A4C116
    K(49) = &H1E376C08
    K(50) = &H2748774C
    K(51) = &H34B0BCB5
    K(52) = &H391C0CB3
    K(53) = &H4ED8AA4A
    K(54) = &H5B9CCA4F
    K(55) = &H682E6FF3
    K(56) = &H748F82EE
    K(57) = &H78A5636F
    K(58) = &H84C87814
    K(59) = &H8CC70208
    K(60) = &H90BEFFFA
    K(61) = &HA4506CEB
    K(62) = &HBEF9A3F7
    K(63) = &HC67178F2
End Sub

'*******************************************************************************
' LShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' A left shift takes all the set binary bits and moves them left, in-filling
' with zeros in the vacated bits on the right. This function is equivalent to
' the << operator in Java and C++
'*******************************************************************************
Private Function LShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    '
    ' A shift of zero is no shift at all.
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
        
    ' A shift of 31 will result in the right most bit becoming the left most
    ' bit and all other bits being cleared
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
        
    ' A shift of less than zero or more than 31 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        err.Raise 6
    End If
    
    ' If the left most bit that remains will end up in the negative bit
    ' position (&H80000000) we would end up with an overflow if we took the
    ' standard route. We need to strip the left most bit and add it back
    ' afterwards.
    If (lValue And m_l2Power(31 - iShiftBits)) Then
    
        ' (Value And OnBits(31 - (Shift + 1))) chops off the left most bits that
        ' we are shifting into, but also the left most bit we still want as this
        ' is going to end up in the negative bit marker position (&H80000000).
        ' After the multiplication/shift we Or the result with &H80000000 to
        ' turn the negative bit on.
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * _
            m_l2Power(iShiftBits)) Or &H80000000
    
    Else
    
        ' (Value And OnBits(31-Shift)) chops off the left most bits that we are
        ' shifting into so we do not get an overflow error when we do the
        ' multiplication/shift
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * _
            m_l2Power(iShiftBits))
        
    End If
End Function

'*******************************************************************************
' RShift (FUNCTION)
'
' PARAMETERS:
' (In) - lValue     - Long    - The value to be shifted
' (In) - iShiftBits - Integer - The number of bits to shift the value by
'
' RETURN VALUE:
' Long - The shifted long integer
'
' DESCRIPTION:
' The right shift of an unsigned long integer involves shifting all the set bits
' to the right and in-filling on the left with zeros. This function is
' equivalent to the >>> operator in Java or the >> operator in C++ when used on
' an unsigned long.
'*******************************************************************************
Private Function RShift(ByVal lValue As Long, _
                        ByVal iShiftBits As Integer) As Long
    
    ' NOTE: If you can guarantee that the Shift parameter will be in the
    ' range 1 to 30 you can safely strip of this first nested if structure for
    ' speed.
    '
    ' A shift of zero is no shift at all
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
        
    ' A shift of 31 will clear all bits and move the left most bit to the right
    ' most bit position
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
        
    ' A shift of less than zero or more than 31 is undefined
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        err.Raise 6
    End If
    
    ' We do not care about the top most bit or the final bit, the top most bit
    ' will be taken into account in the next stage, the final bit (whether it
    ' is an odd number or not) is being shifted into, so we do not give a jot
    ' about it
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    
    ' If the top most bit (&H80000000) was set we need to do things differently
    ' as in a normal VB signed long integer the top most bit is used to indicate
    ' the sign of the number, when it is set it is a negative number, so just
    ' deviding by a factor of 2 as above would not work.
    ' NOTE: (lValue And  &H80000000) is equivalent to (lValue < 0), you could
    ' get a very marginal speed improvement by changing the test to (lValue < 0)
    If (lValue And &H80000000) Then
        ' We take the value computed so far, and then add the left most negative
        ' bit after it has been shifted to the right the appropriate number of
        ' places
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

'*******************************************************************************
' AddUnsigned (FUNCTION)
'
' PARAMETERS:
' (In) - lX - Long - First value
' (In) - lY - Long - Second value
'
' RETURN VALUE:
' Long - Result
'
' DESCRIPTION:
' Adds two potentially large unsigned numbers without overflowing
'*******************************************************************************
Private Function AddUnsigned(ByVal lX As Long, _
                             ByVal lY As Long) As Long
    Dim lX4     As Long
    Dim lY4     As Long
    Dim lX8     As Long
    Dim lY8     As Long
    Dim lResult As Long
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

'*******************************************************************************
' Ch (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Ch(ByVal x As Long, _
                    ByVal y As Long, _
                    ByVal z As Long) As Long
    Ch = ((x And y) Xor ((Not x) And z))
End Function

'*******************************************************************************
' Maj (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Maj(ByVal x As Long, _
                     ByVal y As Long, _
                     ByVal z As Long) As Long
    Maj = ((x And y) Xor (x And z) Xor (y And z))
End Function

'*******************************************************************************
' S (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function (rotate right)
'*******************************************************************************
Private Function S(ByVal x As Long, _
                   ByVal n As Long) As Long
    S = (RShift(x, (n And m_lOnBits(4))) Or LShift(x, (32 - (n And m_lOnBits(4)))))
End Function

'*******************************************************************************
' R (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function (just a right shift)
'*******************************************************************************
Private Function R(ByVal x As Long, _
                   ByVal n As Long) As Long
    R = RShift(x, CInt(n And m_lOnBits(4)))
End Function

'*******************************************************************************
' Sigma0 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Sigma0(ByVal x As Long) As Long
    Sigma0 = (S(x, 2) Xor S(x, 13) Xor S(x, 22))
End Function

'*******************************************************************************
' Sigma1 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Sigma1(ByVal x As Long) As Long
    Sigma1 = (S(x, 6) Xor S(x, 11) Xor S(x, 25))
End Function

'*******************************************************************************
' Gamma0 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Gamma0(ByVal x As Long) As Long
    Gamma0 = (S(x, 7) Xor S(x, 18) Xor R(x, 3))
End Function

'*******************************************************************************
' Gamma1 (FUNCTION)
'
' DESCRIPTION:
' SHA-256 function
'*******************************************************************************
Private Function Gamma1(ByVal x As Long) As Long
    Gamma1 = (S(x, 17) Xor S(x, 19) Xor R(x, 10))
End Function

'*******************************************************************************
' ConvertToWordArray (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - String message
'
' RETURN VALUE:
' Long() - Converted message as long array
'
' DESCRIPTION:
' Takes the string message and puts it in a long array with padding according to
' the SHA-256 rules (similar to MD5 routine).
'*******************************************************************************
Private Function ConvertToWordArray(sMessage As String) As Long()
    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long
    
    Const MODULUS_BITS      As Long = 512
    Const CONGRUENT_BITS    As Long = 448
    
    lMessageLength = Len(sMessage)
    
    ' Get padded number of words. Message needs to be congruent to 448 bits,
    ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
    ' it must still have another 512 bits added. 512 bits = 64 bytes
    ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
    ' be a multiple of 16 (i.e. 16 * 4 (bytes) * 8 (bits))
    lNumberOfWords = (((lMessageLength + _
        ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ _
        (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * _
        (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    ' Combine each block of 4 bytes (ascii code of character) into one long
    ' value and store in the message. The high-order (most significant) bit of
    ' each byte is listed first. However, unlike MD5 we put the high-order
    ' (most significant) byte first in each word.
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        ' Each word is 4 bytes
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        
        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE
        
        ' NOTE: This is where we are using just the first byte of each unicode
        ' character, you may want to make the change here, or to the SHA256 method
        ' so it accepts a byte array.
        lByte = AscB(Mid(sMessage, lByteCount + 1, 1))
        
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
    ' bits stored in the last two words
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    ' Add a terminating 1 bit, all the rest of the bits to the end of the
    ' word array will default to zero
    lWordArray(lWordCount) = lWordArray(lWordCount) Or _
        LShift(&H80, lBytePosition)

    ' We put the length of the message in bits into the last two words, to get
    ' the length in bits we need to multiply by 8 (or left shift 3). This left
    ' shifted value is put in the last word. Any bits shifted off the left edge
    ' need to be put in the penultimate word, we can work out which bits by shifting
    ' right the length by 29 bits.
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

'*******************************************************************************
' SHA256 (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sMessage - String - Message to digest
'
' RETURN VALUE:
' String - The digest
'
' DESCRIPTION:
' Takes a string and uses the SHA-256 digest to produce a signature for it.
'
' NOTE: Due to the way in which the string is processed the routine assumes a
' single byte character set. VB passes unicode (2-byte) character strings, the
' ConvertToWordArray function uses on the first byte for each character. This
' has been done this way for ease of use, to make the routine truely portable
' you could accept a byte array instead, it would then be up to the calling
' routine to make sure that the byte array is generated from their string in
' a manner consistent with the string type.
'*******************************************************************************
Public Function SHA256(sMessage As String, salt As String) As String
    Dim HASH(7) As Long
    Dim M()     As Long
    Dim W(63)   As Long
    Dim a       As Long
    Dim b       As Long
    Dim c       As Long
    Dim d       As Long
    Dim e       As Long
    Dim f       As Long
    Dim g       As Long
    Dim h       As Long
    Dim i       As Long
    Dim j       As Long
    Dim T1      As Long
    Dim T2      As Long
    
    ' Initial hash values
    HASH(0) = &H6A09E667
    HASH(1) = &HBB67AE85
    HASH(2) = &H3C6EF372
    HASH(3) = &HA54FF53A
    HASH(4) = &H510E527F
    HASH(5) = &H9B05688C
    HASH(6) = &H1F83D9AB
    HASH(7) = &H5BE0CD19
    
    ' Preprocessing. Append padding bits and length and convert to words
    M = ConvertToWordArray(salt + sMessage)
    
    ' Main loop
    For i = 0 To UBound(M) Step 16
        a = HASH(0)
        b = HASH(1)
        c = HASH(2)
        d = HASH(3)
        e = HASH(4)
        f = HASH(5)
        g = HASH(6)
        h = HASH(7)
        
        For j = 0 To 63
            If j < 16 Then
                W(j) = M(j + i)
            Else
                W(j) = AddUnsigned(AddUnsigned(AddUnsigned(Gamma1(W(j - 2)), _
                    W(j - 7)), Gamma0(W(j - 15))), W(j - 16))
            End If
                
            T1 = AddUnsigned(AddUnsigned(AddUnsigned(AddUnsigned(h, Sigma1(e)), _
                Ch(e, f, g)), K(j)), W(j))
            T2 = AddUnsigned(Sigma0(a), Maj(a, b, c))
            
            h = g
            g = f
            f = e
            e = AddUnsigned(d, T1)
            d = c
            c = b
            b = a
            a = AddUnsigned(T1, T2)
        Next
        
        HASH(0) = AddUnsigned(a, HASH(0))
        HASH(1) = AddUnsigned(b, HASH(1))
        HASH(2) = AddUnsigned(c, HASH(2))
        HASH(3) = AddUnsigned(d, HASH(3))
        HASH(4) = AddUnsigned(e, HASH(4))
        HASH(5) = AddUnsigned(f, HASH(5))
        HASH(6) = AddUnsigned(g, HASH(6))
        HASH(7) = AddUnsigned(h, HASH(7))
    Next
    
    ' Output the 256 bit digest
    SHA256 = LCase(Right("00000000" & Hex(HASH(0)), 8) & _
        Right("00000000" & Hex(HASH(1)), 8) & _
        Right("00000000" & Hex(HASH(2)), 8) & _
        Right("00000000" & Hex(HASH(3)), 8) & _
        Right("00000000" & Hex(HASH(4)), 8) & _
        Right("00000000" & Hex(HASH(5)), 8) & _
        Right("00000000" & Hex(HASH(6)), 8) & _
        Right("00000000" & Hex(HASH(7)), 8))
End Function

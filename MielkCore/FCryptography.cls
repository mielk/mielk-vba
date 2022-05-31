VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FCryptography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const CLASS_NAME As String = "FCryptography"
    '----------------------------------------------------------------------------------------------------------

Private Type FourBytes
    a As Byte
    b As Byte
    C As Byte
    d As Byte
End Type

Private Type OneLong
    l As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

    '----------------------------------------------------------------------------------------------------------
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As LongPtr
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
#End If
    '----------------------------------------------------------------------------------------------------------




    '----------------------------------------------------------------------------------------------------------
' Description:  Code for generating UUID without using Scriplet.TypeLib.
' Author:       Gus Molina
' Date:         2000
    '----------------------------------------------------------------------------------------------------------
Public Function createUUID() As String
    Dim udtGUID As GUID
    '------------------------------------------------------------------------------------------------------
    
    If (CoCreateGuid(udtGUID) = 0) Then
        createUUID = _
                VBA.String(8 - VBA.Len(VBA.Hex$(udtGUID.Data1)), "0") & VBA.Hex$(udtGUID.Data1) & "-" & _
                VBA.String(4 - VBA.Len(VBA.Hex$(udtGUID.Data2)), "0") & VBA.Hex$(udtGUID.Data2) & "-" & _
                VBA.String(4 - VBA.Len(VBA.Hex$(udtGUID.Data3)), "0") & VBA.Hex$(udtGUID.Data3) & "-" & _
                VBA.IIf((udtGUID.Data4(0) < &H10), "0", "") & VBA.Hex$(udtGUID.Data4(0)) & _
                VBA.IIf(udtGUID.Data4(1) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(1)) & _
                "-" & _
                VBA.IIf(udtGUID.Data4(2) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(2)) & _
                VBA.IIf(udtGUID.Data4(3) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(3)) & _
                VBA.IIf(udtGUID.Data4(4) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(4)) & _
                VBA.IIf(udtGUID.Data4(5) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(5)) & _
                VBA.IIf(udtGUID.Data4(6) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(6)) & _
                VBA.IIf(udtGUID.Data4(7) < &H10, "0", "") & VBA.Hex$(udtGUID.Data4(7))
    End If

End Function



' Convert the string into bytes so we can use the above Functions
' From Chris Hulbert: http://splinter.com.au/blog
Public Function SHA1HASH(str)
    Dim i As Integer
    Dim arr() As Byte
    '------------------------------------------------------------------------------------------------------
    
    ReDim arr(0 To VBA.Len(str) - 1)
    For i = 0 To VBA.Len(str) - 1
        arr(i) = VBA.asc(VBA.mid(str, i + 1, 1))
    Next i
    SHA1HASH = VBA.Replace(VBA.LCase(HexDefaultSHA1(arr)), " ", "")
    
End Function



Private Function HexDefaultSHA1(message() As Byte) As String
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
    DefaultSHA1 message, H1, H2, H3, H4, H5
    HexDefaultSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Private Function HexSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long) As String
    Dim H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long
    xSHA1 message, Key1, Key2, Key3, Key4, H1, H2, H3, H4, H5
    HexSHA1 = DecToHex5(H1, H2, H3, H4, H5)
End Function

Private Sub DefaultSHA1(message() As Byte, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
    xSHA1 message, &H5A827999, &H6ED9EBA1, &H8F1BBCDC, &HCA62C1D6, H1, H2, H3, H4, H5
End Sub

Private Sub xSHA1(message() As Byte, ByVal Key1 As Long, ByVal Key2 As Long, ByVal Key3 As Long, ByVal Key4 As Long, H1 As Long, H2 As Long, H3 As Long, H4 As Long, H5 As Long)
    Dim u As Long, p As Long
    Dim fb As FourBytes, OL As OneLong
    Dim i As Integer
    Dim w(0 To 80) As Long
    Dim a As Long, b As Long, C As Long, d As Long, E As Long
    Dim t As Long
    
    H1 = &H67452301: H2 = &HEFCDAB89: H3 = &H98BADCFE: H4 = &H10325476: H5 = &HC3D2E1F0
    
    u = UBound(message) + 1: OL.l = U32ShiftLeft3(u): a = u \ &H20000000: LSet fb = OL 'U32ShiftRight29(U)
    
    ReDim Preserve message(0 To (u + 8 And -64) + 63)
    message(u) = 128
    
    u = UBound(message)
    message(u - 4) = a
    message(u - 3) = fb.d
    message(u - 2) = fb.C
    message(u - 1) = fb.b
    message(u) = fb.a
    
    While p < u
        For i = 0 To 15
            fb.d = message(p)
            fb.C = message(p + 1)
            fb.b = message(p + 2)
            fb.a = message(p + 3)
            LSet OL = fb
            w(i) = OL.l
            p = p + 4
        Next i
    
        For i = 16 To 79
            w(i) = U32RotateLeft1(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16))
        Next i
    
        a = H1: b = H2: C = H3: d = H4: E = H5
    
        For i = 0 To 19
            t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key1), ((b And C) Or ((Not b) And d)))
            E = d: d = C: C = U32RotateLeft30(b): b = a: a = t
        Next i
        For i = 20 To 39
            t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key2), (b Xor C Xor d))
            E = d: d = C: C = U32RotateLeft30(b): b = a: a = t
        Next i
        For i = 40 To 59
            t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key3), ((b And C) Or (b And d) Or (C And d)))
            E = d: d = C: C = U32RotateLeft30(b): b = a: a = t
        Next i
        For i = 60 To 79
            t = U32Add(U32Add(U32Add(U32Add(U32RotateLeft5(a), E), w(i)), Key4), (b Xor C Xor d))
            E = d: d = C: C = U32RotateLeft30(b): b = a: a = t
        Next i
    
        H1 = U32Add(H1, a): H2 = U32Add(H2, b): H3 = U32Add(H3, C): H4 = U32Add(H4, d): H5 = U32Add(H5, E)
    Wend
End Sub

Private Function U32Add(ByVal a As Long, ByVal b As Long) As Long
    If (a Xor b) < 0 Then
        U32Add = a + b
    Else
        U32Add = (a Xor &H80000000) + b Xor &H80000000
    End If
End Function

Private Function U32ShiftLeft3(ByVal a As Long) As Long
    U32ShiftLeft3 = (a And &HFFFFFFF) * 8
    If a And &H10000000 Then U32ShiftLeft3 = U32ShiftLeft3 Or &H80000000
End Function

Private Function U32ShiftRight29(ByVal a As Long) As Long
    U32ShiftRight29 = (a And &HE0000000) \ &H20000000 And 7
End Function

Private Function U32RotateLeft1(ByVal a As Long) As Long
    U32RotateLeft1 = (a And &H3FFFFFFF) * 2
    If a And &H40000000 Then U32RotateLeft1 = U32RotateLeft1 Or &H80000000
    If a And &H80000000 Then U32RotateLeft1 = U32RotateLeft1 Or 1
End Function

Private Function U32RotateLeft5(ByVal a As Long) As Long
    U32RotateLeft5 = (a And &H3FFFFFF) * 32 Or (a And &HF8000000) \ &H8000000 And 31
    If a And &H4000000 Then U32RotateLeft5 = U32RotateLeft5 Or &H80000000
End Function

Private Function U32RotateLeft30(ByVal a As Long) As Long
    U32RotateLeft30 = (a And 1) * &H40000000 Or (a And &HFFFC) \ 4 And &H3FFFFFFF
    If a And 2 Then U32RotateLeft30 = U32RotateLeft30 Or &H80000000
End Function

Private Function DecToHex5(ByVal H1 As Long, ByVal H2 As Long, ByVal H3 As Long, ByVal H4 As Long, ByVal H5 As Long) As String
    Dim H As String, l As Long
    DecToHex5 = "00000000 00000000 00000000 00000000 00000000"
    H = Hex(H1): l = Len(H): Mid(DecToHex5, 9 - l, l) = H
    H = Hex(H2): l = Len(H): Mid(DecToHex5, 18 - l, l) = H
    H = Hex(H3): l = Len(H): Mid(DecToHex5, 27 - l, l) = H
    H = Hex(H4): l = Len(H): Mid(DecToHex5, 36 - l, l) = H
    H = Hex(H5): l = Len(H): Mid(DecToHex5, 45 - l, l) = H
End Function

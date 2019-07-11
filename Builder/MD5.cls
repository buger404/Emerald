VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Visual Basic MD5 Implementation
' Robert Hubley and David Midkiff
' modify by simonyan, Support chinese
' Standard MD5 implementation optimised for the Visual Basic environment.
' Conforms to all standards and can be used in digital signature or password
' protection related schemes.
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private State(4) As Long
Private ByteCounter As Long
Private ByteBuffer(63) As Byte
Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21
Property Get RegisterA() As String
RegisterA = State(1)
End Property
Property Get RegisterB() As String
RegisterB = State(2)
End Property
Property Get RegisterC() As String
RegisterC = State(3)
End Property
Property Get RegisterD() As String
RegisterD = State(4)
End Property
Public Function Md5_String_Calc(SourceString As String) As String
MD5Init
MD5Update LenB(StrConv(SourceString, vbFromUnicode)), StringToArray(SourceString)
MD5Final
Md5_String_Calc = GetValues
End Function
Public Function Md5_File_Calc(InFile As String) As String
On Error GoTo errorhandler
GoSub begin
errorhandler:
Md5_File_Calc = ""
Exit Function
begin:
Dim FileO As Integer
FileO = FreeFile
Call FileLen(InFile)
Open InFile For Binary Access Read As #FileO
MD5Init
Do While Not EOF(FileO)
Get #FileO, , ByteBuffer
If Loc(FileO) < LOF(FileO) Then
ByteCounter = ByteCounter + 64
MD5Transform ByteBuffer
End If
Loop
ByteCounter = ByteCounter + (LOF(FileO) Mod 64)
Close #FileO
MD5Final
Md5_File_Calc = GetValues
End Function
Private Function StringToArray(InString As String) As Byte()
Dim I As Integer, bytBuffer() As Byte
ReDim bytBuffer(LenB(StrConv(InString, vbFromUnicode)))
bytBuffer = StrConv(InString, vbFromUnicode)
StringToArray = bytBuffer
End Function
Public Function GetValues() As String
GetValues = LongToString(State(1)) & LongToString(State(2)) & LongToString(State(3)) & LongToString(State(4))
End Function
Private Function LongToString(Num As Long) As String
Dim A As Byte, B As Byte, C As Byte, D As Byte
A = Num And &HFF&
If A < 16 Then LongToString = "0" & Hex(A) Else LongToString = Hex(A)
B = (Num And &HFF00&) \ 256
If B < 16 Then LongToString = LongToString & "0" & Hex(B) Else LongToString = LongToString & Hex(B)
C = (Num And &HFF0000) \ 65536
If C < 16 Then LongToString = LongToString & "0" & Hex(C) Else LongToString = LongToString & Hex(C)
If Num < 0 Then D = ((Num And &H7F000000) \ 16777216) Or &H80& Else D = (Num And &HFF000000) \ 16777216
If D < 16 Then LongToString = LongToString & "0" & Hex(D) Else LongToString = LongToString & Hex(D)
End Function
Public Sub MD5Init()
ByteCounter = 0
State(1) = UnsignedToLong(1732584193#)
State(2) = UnsignedToLong(4023233417#)
State(3) = UnsignedToLong(2562383102#)
State(4) = UnsignedToLong(271733878#)
End Sub
Public Sub MD5Final()
Dim dblBits As Double, padding(72) As Byte, lngBytesBuffered As Long
padding(0) = &H80
dblBits = ByteCounter * 8
lngBytesBuffered = ByteCounter Mod 64
If lngBytesBuffered <= 56 Then MD5Update 56 - lngBytesBuffered, padding Else MD5Update 120 - ByteCounter, padding
padding(0) = UnsignedToLong(dblBits) And &HFF&
padding(1) = UnsignedToLong(dblBits) \ 256 And &HFF&
padding(2) = UnsignedToLong(dblBits) \ 65536 And &HFF&
padding(3) = UnsignedToLong(dblBits) \ 16777216 And &HFF&
padding(4) = 0
padding(5) = 0
padding(6) = 0
padding(7) = 0
MD5Update 8, padding
End Sub
Public Sub MD5Update(InputLen As Long, InputBuffer() As Byte)
Dim II As Integer, I As Integer, J As Integer, K As Integer, lngBufferedBytes As Long, lngBufferRemaining As Long, lngRem As Long
lngBufferedBytes = ByteCounter Mod 64
lngBufferRemaining = 64 - lngBufferedBytes
ByteCounter = ByteCounter + InputLen
If InputLen >= lngBufferRemaining Then
For II = 0 To lngBufferRemaining - 1
ByteBuffer(lngBufferedBytes + II) = InputBuffer(II)
Next II
MD5Transform ByteBuffer
lngRem = (InputLen) Mod 64
For I = lngBufferRemaining To InputLen - II - lngRem Step 64
For J = 0 To 63
ByteBuffer(J) = InputBuffer(I + J)
Next J
MD5Transform ByteBuffer
Next I
lngBufferedBytes = 0
Else
I = 0
End If
For K = 0 To InputLen - I - 1
ByteBuffer(lngBufferedBytes + K) = InputBuffer(I + K)
Next K
End Sub
Private Sub MD5Transform(Buffer() As Byte)
Dim X(16) As Long, A As Long, B As Long, C As Long, D As Long
A = State(1)
B = State(2)
C = State(3)
D = State(4)
Decode 64, X, Buffer
FF A, B, C, D, X(0), S11, -680876936
FF D, A, B, C, X(1), S12, -389564586
FF C, D, A, B, X(2), S13, 606105819
FF B, C, D, A, X(3), S14, -1044525330
FF A, B, C, D, X(4), S11, -176418897
FF D, A, B, C, X(5), S12, 1200080426
FF C, D, A, B, X(6), S13, -1473231341
FF B, C, D, A, X(7), S14, -45705983
FF A, B, C, D, X(8), S11, 1770035416
FF D, A, B, C, X(9), S12, -1958414417
FF C, D, A, B, X(10), S13, -42063
FF B, C, D, A, X(11), S14, -1990404162
FF A, B, C, D, X(12), S11, 1804603682
FF D, A, B, C, X(13), S12, -40341101
FF C, D, A, B, X(14), S13, -1502002290
FF B, C, D, A, X(15), S14, 1236535329
GG A, B, C, D, X(1), S21, -165796510
GG D, A, B, C, X(6), S22, -1069501632
GG C, D, A, B, X(11), S23, 643717713
GG B, C, D, A, X(0), S24, -373897302
GG A, B, C, D, X(5), S21, -701558691
GG D, A, B, C, X(10), S22, 38016083
GG C, D, A, B, X(15), S23, -660478335
GG B, C, D, A, X(4), S24, -405537848
GG A, B, C, D, X(9), S21, 568446438
GG D, A, B, C, X(14), S22, -1019803690
GG C, D, A, B, X(3), S23, -187363961
GG B, C, D, A, X(8), S24, 1163531501
GG A, B, C, D, X(13), S21, -1444681467
GG D, A, B, C, X(2), S22, -51403784
GG C, D, A, B, X(7), S23, 1735328473
GG B, C, D, A, X(12), S24, -1926607734
HH A, B, C, D, X(5), S31, -378558
HH D, A, B, C, X(8), S32, -2022574463
HH C, D, A, B, X(11), S33, 1839030562
HH B, C, D, A, X(14), S34, -35309556
HH A, B, C, D, X(1), S31, -1530992060
HH D, A, B, C, X(4), S32, 1272893353
HH C, D, A, B, X(7), S33, -155497632
HH B, C, D, A, X(10), S34, -1094730640
HH A, B, C, D, X(13), S31, 681279174
HH D, A, B, C, X(0), S32, -358537222
HH C, D, A, B, X(3), S33, -722521979
HH B, C, D, A, X(6), S34, 76029189
HH A, B, C, D, X(9), S31, -640364487
HH D, A, B, C, X(12), S32, -421815835
HH C, D, A, B, X(15), S33, 530742520
HH B, C, D, A, X(2), S34, -995338651
II A, B, C, D, X(0), S41, -198630844
II D, A, B, C, X(7), S42, 1126891415
II C, D, A, B, X(14), S43, -1416354905
II B, C, D, A, X(5), S44, -57434055
II A, B, C, D, X(12), S41, 1700485571
II D, A, B, C, X(3), S42, -1894986606
II C, D, A, B, X(10), S43, -1051523
II B, C, D, A, X(1), S44, -2054922799
II A, B, C, D, X(8), S41, 1873313359
II D, A, B, C, X(15), S42, -30611744
II C, D, A, B, X(6), S43, -1560198380
II B, C, D, A, X(13), S44, 1309151649
II A, B, C, D, X(4), S41, -145523070
II D, A, B, C, X(11), S42, -1120210379
II C, D, A, B, X(2), S43, 718787259
II B, C, D, A, X(9), S44, -343485551
State(1) = LongOverflowAdd(State(1), A)
State(2) = LongOverflowAdd(State(2), B)
State(3) = LongOverflowAdd(State(3), C)
State(4) = LongOverflowAdd(State(4), D)
End Sub
Private Sub Decode(Length As Integer, OutputBuffer() As Long, InputBuffer() As Byte)
Dim intDblIndex As Integer, intByteIndex As Integer, dblSum As Double
For intByteIndex = 0 To Length - 1 Step 4
dblSum = InputBuffer(intByteIndex) + InputBuffer(intByteIndex + 1) * 256# + InputBuffer(intByteIndex + 2) * 65536# + InputBuffer(intByteIndex + 3) * 16777216#
OutputBuffer(intDblIndex) = UnsignedToLong(dblSum)
intDblIndex = intDblIndex + 1
Next intByteIndex
End Sub
Private Function FF(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
A = LongOverflowAdd4(A, (B And C) Or (Not (B) And D), X, ac)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)
End Function
Private Function GG(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
A = LongOverflowAdd4(A, (B And D) Or (C And Not (D)), X, ac)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)
End Function
Private Function HH(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
A = LongOverflowAdd4(A, B Xor C Xor D, X, ac)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)
End Function
Private Function II(A As Long, B As Long, C As Long, D As Long, X As Long, S As Long, ac As Long) As Long
A = LongOverflowAdd4(A, C Xor (B Or Not (D)), X, ac)
A = LongLeftRotate(A, S)
A = LongOverflowAdd(A, B)
End Function
Function LongLeftRotate(value As Long, Bits As Long) As Long
Dim lngSign As Long, lngI As Long
Bits = Bits Mod 32
If Bits = 0 Then LongLeftRotate = value: Exit Function
For lngI = 1 To Bits
lngSign = value And &HC0000000
value = (value And &H3FFFFFFF) * 2
value = value Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
Next
LongLeftRotate = value
End Function
Private Function LongOverflowAdd(Val1 As Long, Val2 As Long) As Long
Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&)
lngOverflow = lngLowWord \ 65536
lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
LongOverflowAdd = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function
Private Function LongOverflowAdd4(Val1 As Long, Val2 As Long, val3 As Long, val4 As Long) As Long
Dim lngHighWord As Long, lngLowWord As Long, lngOverflow As Long
lngLowWord = (Val1 And &HFFFF&) + (Val2 And &HFFFF&) + (val3 And &HFFFF&) + (val4 And &HFFFF&)
lngOverflow = lngLowWord \ 65536
lngHighWord = (((Val1 And &HFFFF0000) \ 65536) + ((Val2 And &HFFFF0000) \ 65536) + ((val3 And &HFFFF0000) \ 65536) + ((val4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
LongOverflowAdd4 = UnsignedToLong((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))
End Function
Private Function UnsignedToLong(value As Double) As Long
If value < 0 Or value >= OFFSET_4 Then Error 6
If value <= MAXINT_4 Then UnsignedToLong = value Else UnsignedToLong = value - OFFSET_4
End Function
Private Function LongToUnsigned(value As Long) As Double
If value < 0 Then LongToUnsigned = value + OFFSET_4 Else LongToUnsigned = value
End Function

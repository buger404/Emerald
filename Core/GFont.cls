VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald 相关代码

Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Dim brush As Long, fFile As String, fFam As Long, StrF(2) As Long, Font(8, 1 To 100) As Long
Public EMHeight As Long
'========================================================
'   Init
    Public Sub MakeFont(Name As String)
        GdipCreateFontFamilyFromName StrPtr(Name), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("微软雅黑"), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("宋体"), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("System"), 0, fFam
        
        Dim Flag As Long
        For I = 0 To 2
            GdipCreateStringFormat 0, 0, StrF(I)
            GdipSetStringFormatAlign StrF(I), I
            'GdipStringFormatGetGenericTypographic strF(i)
            GdipGetStringFormatFlags StrF(I), Flag
            GdipSetStringFormatFlags StrF(I), (Flag Or StringFormatFlagsMeasureTrailingSpaces)
        Next
        PoolCreateSolidFill argb(255, 64, 64, 64), brush
        
        EmeraldInstalled = True
    End Sub
    Public Sub Dispose()
        GdipDeleteFontFamily fFam
        If fFile <> "" Then RemoveFontResource fFile
        Exit Sub
        PoolDeleteBrush brush
        For I = 0 To 2
            PoolDeleteStringFormat StrF(I)
        Next
        For S = 0 To 8
            For I = 1 To 100
                If Font(S, I) <> 0 Then PoolDeleteFont Font(S, I)
            Next
       Next
    End Sub
'========================================================
'   FontFromFile
    Public Sub AddFont(ByVal File As String)
        fFile = File
        AddFontResource File
    End Sub
'========================================================
'   Writer
    Public Function GetWidth(g As Long, Text As String, Size As Long, ByVal align As StringAlignment, ByVal style As FontStyle) As Long
        Dim r As RECTF, r2 As RECTF
        GdipMeasureString g, StrPtr(Text), Len(Text), Font(style, Size), r2, StrF(align), r, 0, 0
        
        GetWidth = r.Right
    End Function
    Public Function GetHeight(g As Long, Text As String, Size As Long, ByVal align As StringAlignment, ByVal style As FontStyle) As Long
        Dim r As RECTF, r2 As RECTF
        GdipMeasureString g, StrPtr(Text), Len(Text), Font(style, Size), r2, StrF(align), r, 0, 0
        
        GetHeight = r.Bottom
    End Function
    Public Sub Writes(Text As String, X As Long, y As Long, g As Long, Color As Long, Size As Long, w As Long, h As Long, ByVal align As StringAlignment, ByVal style As FontStyle)
        On Error Resume Next
        If Font(style, Size) = 0 Then
            PoolCreateFont fFam, Size, style, UnitPixel, Font(style, Size)
        End If
        If EMHeight = 0 Then EMHeight = GetHeight(g, "ABCDEFG", 18, StringAlignmentNear, FontStyleRegular)
        
        If Color <> 0 Then GdipSetSolidFillColor brush, Color
        GdipDrawString g, StrPtr(Text), -1, Font(style, Size), NewRectF(X, y, w, h), StrF(align), brush
    End Sub
'========================================================

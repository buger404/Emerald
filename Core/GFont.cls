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
'========================================================
'   Init
    Public Sub MakeFont(name As String)
        GdipCreateFontFamilyFromName StrPtr(name), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("微软雅黑"), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("宋体"), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("System"), 0, fFam
        
        Dim Flag As Long
        For i = 0 To 2
            GdipCreateStringFormat 0, 0, StrF(i)
            GdipSetStringFormatAlign StrF(i), i
            'GdipStringFormatGetGenericTypographic strF(i)
            GdipGetStringFormatFlags StrF(i), Flag
            GdipSetStringFormatFlags StrF(i), (Flag Or StringFormatFlagsMeasureTrailingSpaces)
        Next
        PoolCreateSolidFill argb(255, 64, 64, 64), brush
        
        EmeraldInstalled = True
    End Sub
    Public Sub Dispose()
        GdipDeleteFontFamily fFam
        If fFile <> "" Then RemoveFontResource fFile
        Exit Sub
        PoolDeleteBrush brush
        For i = 0 To 2
            PoolDeleteStringFormat StrF(i)
        Next
        For s = 0 To 8
            For i = 1 To 100
                If Font(s, i) <> 0 Then PoolDeleteFont Font(s, i)
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
    Public Function GetWidth(g As Long, text As String, size As Long, ByVal align As StringAlignment, ByVal style As FontStyle) As Long
        Dim r As RECTF, r2 As RECTF
        GdipMeasureString g, StrPtr(text), Len(text), Font(style, size), r2, StrF(align), r, 0, 0
        
        GetWidth = r.Right
    End Function
    Public Function GetHeight(g As Long, text As String, size As Long, ByVal align As StringAlignment, ByVal style As FontStyle) As Long
        Dim r As RECTF, r2 As RECTF
        GdipMeasureString g, StrPtr(text), Len(text), Font(style, size), r2, StrF(align), r, 0, 0
        
        GetHeight = r.Bottom
    End Function
    Public Sub Writes(text As String, x As Long, y As Long, g As Long, Color As Long, size As Long, w As Long, h As Long, ByVal align As StringAlignment, ByVal style As FontStyle)
        On Error Resume Next
        If Font(style, size) = 0 Then
            PoolCreateFont fFam, size, style, UnitPixel, Font(style, size)
        End If
        
        If Color <> 0 Then GdipSetSolidFillColor brush, Color
        GdipDrawString g, StrPtr(text), -1, Font(style, size), NewRectF(x, y, w, h), StrF(align), brush
    End Sub
'========================================================

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
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Dim brush As Long, fFile As String, fFam As Long, strF(2) As Long, Font(2, 1 To 100) As Long
'========================================================
'   Init
    Public Sub MakeFont(name As String)
        GdipCreateFontFamilyFromName StrPtr(name), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("Î¢ÈíÑÅºÚ"), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("ËÎÌå"), 0, fFam
        If fFam = 0 Then GdipCreateFontFamilyFromName StrPtr("System"), 0, fFam
        
        For i = 0 To 2
            GdipCreateStringFormat 0, 0, strF(i)
            GdipSetStringFormatAlign strF(i), i
        Next
        GdipCreateSolidFill argb(255, 64, 64, 64), brush
    End Sub
    Public Sub Dispose()
        If fFile <> "" Then RemoveFontResource fFile
        GdipDeleteBrush brush: GdipDeleteFontFamily fFam
        For i = 0 To 2
            GdipDeleteStringFormat strF(i)
        Next
        For s = 0 To 2
            For i = 1 To 100
                If Font(s, i) <> 0 Then GdipDeleteFont Font(s, i)
            Next
       Next
    End Sub
'========================================================
'   FontFromFile
    Public Sub AddFont(ByVal file As String)
        fFile = file
        AddFontResource file
    End Sub
'========================================================
'   Writer
    Public Sub Writes(text As String, X As Long, Y As Long, g As Long, Color As Long, size As Long, w As Long, h As Long, ByVal align As StringAlignment, ByVal style As FontStyle)
        On Error Resume Next
        If Font(style, size) = 0 Then
            If style = 0 Then GdipCreateFont fFam, size, FontStyle.FontStyleRegular, UnitPixel, Font(style, size)
            If style = 1 Then GdipCreateFont fFam, size, FontStyle.FontStyleBold, UnitPixel, Font(style, size)
            If style = 2 Then GdipCreateFont fFam, size, FontStyle.FontStyleItalic, UnitPixel, Font(style, size)
        End If
        
        If Color <> 0 Then GdipSetSolidFillColor brush, Color
        GdipDrawString g, StrPtr(text), -1, Font(style, size), NewRectF(X, Y, w, h), strF(align), brush
    End Sub
'========================================================

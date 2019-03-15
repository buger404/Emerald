VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
Private Type GMem
    kind As Integer
    hwnd As Long
    name As String
    folder As String
    w As Long
    h As Long
End Type
Private Type GDA
    x As Long
    y As Long
    w As Long
    h As Long
End Type
Dim Mems() As GMem, DrawBox As Object, brush As Long, path As Long, Pen As Long
Public CDC As Long, gg As Long
'========================================================
'   Init
    Public Sub Create(dBox As Object)
        ReDim Mems(0)
        
        CDC = CreateCDC(GW, GH)
        NewGM CDC, 0, ""
        GdipCreateFromHDC CDC, gg
        GdipSetSmoothingMode gg, SmoothingModeAntiAlias
        GdipSetTextRenderingHint gg, TextRenderingHintAntiAlias
        
        Set DrawBox = dBox
        
        GdipCreateSolidFill 0, brush
        GdipCreatePen1 0, 1, UnitPixel, Pen
        GdipCreatePath FillModeWinding, path
    End Sub
    Public Sub Dispose()
        GdipDeleteBrush brush
        GdipDeletePen Pen
        GdipDeletePath path
        For i = 1 To UBound(Mems)
            Select Case Mems(i).kind
                Case 0 'DC
                    DeleteObject Mems(i).hwnd
                Case 1 'Graphics
                    GdipDeleteGraphics Mems(i).hwnd
            End Select
        Next
    End Sub
'========================================================
'   Print
    Public Sub Clear()
        GdipGraphicsClear gg, argb(255, 255, 255, 255)
    End Sub
    Public Sub Update()
        CallByName DrawBox, "Update", VbMethod
    End Sub
    Public Sub Display(DC As Long)
        BitBlt DC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
    End Sub
'========================================================
'   RunTime
    Private Sub NewGM(mem As Long, kind As Integer, name As String, Optional w, Optional h, Optional folder As String = "")
        ReDim Preserve Mems(UBound(Mems) + 1)
        With Mems(UBound(Mems))
            .hwnd = mem
            .kind = kind
            .name = name
            .folder = folder
            If Not IsMissing(w) Then .w = w
            If Not IsMissing(h) Then .h = h
        End With
    End Sub
'========================================================
'   Images
    Public Sub NewImages(path As String, Optional arg1, Optional arg2)
        Dim f As String
        If Right(path, 1) <> "\" Then path = path & "\"
        
        f = Dir(path)
        Do While f <> ""
            If IsMissing(arg1) Then
                newImage path & f
            ElseIf IsMissing(arg2) Then
                newImage path & f, arg1
            Else
                newImage path & f, arg1, arg2
            End If
            f = Dir()
        Loop
    End Sub
    Public Sub newImage(path As String, Optional arg1, Optional arg2)
        'when arg1 is missing ; i.w = i.w
        'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
        'other ; i.w = arg1
        Dim si As Integer, i As Long, w As Long, h As Long, DC As Long, g As Long
        If IsMissing(arg1) Then si = 0
        If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
        If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
        
        GdipCreateBitmapFromFile StrPtr(path), i
        GdipGetImageWidth i, w: GdipGetImageHeight i, h
        
        If si = 1 Then w = Round(w * arg1): h = Round(h * arg1)
        If si = 2 Then w = arg1: h = arg2
        
        DC = CreateCDC(w, h)
        GdipCreateFromHDC DC, g
        GdipDrawImageRect g, i, 0, 0, w, h
        
        GdipDisposeImage i: GdipDeleteGraphics g
        
        Dim temp() As String
        temp = Split(path, "\")
        NewGM DC, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
    End Sub
    Public Sub DrawImage(n, x As Long, y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional alpha, Optional Pos As Integer = 0)
        Dim b As BLENDFUNCTION, Index As Integer, bl As Long
        
        If Not IsMissing(alpha) Then
            If alpha < 0 Then alpha = 0
            If alpha > 1 Then alpha = 1
            With b
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(alpha * 255)
            End With
            CopyMemory bl, b, 4
        End If
        
        If TypeName(n) = "String" Then
            Index = GetImage(n)
        Else
            Index = n
        End If
        
        If IsMissing(cx) Then cx = 0
        If IsMissing(cy) Then cy = 0
        If IsMissing(cw) Then cw = Mems(Index).w - cx
        If IsMissing(ch) Then ch = Mems(Index).h - cy
        
        If Pos = 1 Then x = Int(x - cw / 2): y = Int(y - ch / 2)
        If Pos = 2 Then x = x - cw
        If Pos = 3 Then y = y - ch
        If Pos = 4 Then x = x + cw
        If Pos = 5 Then y = y + ch
        
        If IsMissing(alpha) Then
            BitBlt CDC, x, y, cw, ch, Mems(Index).hwnd, cx, cy, vbSrcCopy
        Else
            AlphaBlend CDC, x, y, cw, ch, Mems(Index).hwnd, cx, cy, cw, ch, bl
        End If
        
        With DrawF
            .Left = x
            .top = y
            .Right = cw
            .Bottom = ch
        End With
    End Sub
    Private Function GetImage(ByVal name As String) As Integer
        For i = 1 To UBound(Mems)
            If Mems(i).kind = 0 Then
                If Mems(i).name = name Then GetImage = i: Exit For
            End If
        Next
    End Function
'========================================================
'   Writer
    Public Sub Writes(ByVal text As String, x As Long, y As Long, Optional size As Long = 14, Optional color As Long, Optional w As Long = 0, Optional h As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular)
        EF.Writes text, x, y, gg, color, size, w, h, align, style
        With DrawF
            .Left = x
            .top = y
            .Right = w
            .Bottom = h
        End With
    End Sub
'========================================================
'   Shape
    Public Sub Paint(ByVal shape As Integer, x As Long, y As Long, w As Long, h As Long, Optional color As Long, Optional Radius As Long, Optional style As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        GdipResetPath path
        If shape = 0 Then GdipAddPathRectangle path, x - 1, y - 1, w, h
        If shape = 1 Then GdipAddPathEllipse path, x, y, w - 1, h - 1
        
        If color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, color
            If style = 1 Then GdipSetPenColor Pen, color
        End If
        
        If style = 0 Then GdipFillPath gg, brush, path
        If style = 1 Then GdipDrawPath gg, Pen, path
        
        With DrawF
            .Left = x
            .top = y
            .Right = w
            .Bottom = h
        End With
    End Sub
'========================================================

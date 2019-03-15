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
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As clsid) As Long
Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
Private Type GGIF
    time As Long
    frames() As Long
    tick As Long
    count As Long
End Type
Private Type GMem
    GIF As GGIF
    kind As Integer
    hwnd As Long
    name As String
    folder As String
    W As Long
    H As Long
End Type
Private Type GDA
    X As Long
    Y As Long
    W As Long
    H As Long
End Type
Private Const GTick As Long = 1000 / 30
Dim Mems() As GMem, DrawBox As Object, brush As Long, path As Long, Pen As Long
Dim OutOfScroll As Boolean, ScrollMode As Boolean, ScrollX As Long, ScrollY As Long, ScrollBX As Long, ScrollBY As Long, CRgn As Long
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
            If Mems(i).GIF.count > 0 Then
                For s = 0 To Mems(i).GIF.count
                    DeleteObject Mems(i).GIF.frames(s)
                Next
            End If
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
    Public Sub Display(dc As Long)
        BitBlt dc, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
    End Sub
'========================================================
'   RunTime
    Private Sub NewGM(mem As Long, kind As Integer, name As String, Optional W, Optional H, Optional folder As String = "")
        ReDim Preserve Mems(UBound(Mems) + 1)
        With Mems(UBound(Mems))
            .hwnd = mem
            .kind = kind
            .name = name
            .folder = folder
            If Not IsMissing(W) Then .W = W
            If Not IsMissing(H) Then .H = H
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
        Dim si As Integer, i As Long, W As Long, H As Long, dc As Long, g As Long
        If IsMissing(arg1) Then si = 0
        If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
        If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
        
        GdipCreateBitmapFromFile StrPtr(path), i
        GdipGetImageWidth i, W: GdipGetImageHeight i, H
        
        If si = 1 Then W = Round(W * arg1): H = Round(H * arg1)
        If si = 2 Then W = arg1: H = arg2
        
        Dim gC As clsid, gL As Long, index As Long, temp() As String
        CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
        GdipImageGetFrameCount i, gC, gL
        If gL > 0 Then  '¶¯Ì¬GIFÖ§³Ö
            temp = Split(path, "\")
            NewGM 0, 0, temp(UBound(temp)), W, H, temp(UBound(temp) - 1)
            With Mems(UBound(Mems))
                .GIF.count = gL
                ReDim .GIF.frames(gL)
                For index = 0 To gL
                    GdipImageSelectActiveFrame i, gC, index
                    dc = CreateCDC(W, H)
                    GdipCreateFromHDC dc, g
                    GdipDrawImageRect g, i, 0, 0, W, H
                    GdipDeleteGraphics g
                    .GIF.frames(index) = dc
                Next
            End With
            
            GdipDisposeImage i
            Exit Sub
        End If
        
        dc = CreateCDC(W, H)
        GdipCreateFromHDC dc, g
        GdipDrawImageRect g, i, 0, 0, W, H
        
        GdipDisposeImage i: GdipDeleteGraphics g
        
        temp = Split(path, "\")
        NewGM dc, 0, temp(UBound(temp)), W, H, temp(UBound(temp) - 1)
    End Sub
    Public Sub DrawImage(n, X As Long, Y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional alpha, Optional Pos As Integer = 0)
        If OutOfScroll Then Exit Sub
        
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        
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
            index = GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(cx) Then cx = 0
        If IsMissing(cy) Then cy = 0
        If IsMissing(cw) Then cw = Mems(index).W - cx
        If IsMissing(ch) Then ch = Mems(index).H - cy
        
        If Pos = 1 Then X = Int(X - cw / 2): Y = Int(Y - ch / 2)
        If Pos = 2 Then X = X - cw
        If Pos = 3 Then Y = Y - ch
        If Pos = 4 Then X = X + cw
        If Pos = 5 Then Y = Y + ch
        
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim ret As Long, Crect As RECT
            With Crect
                .Left = X: .top = Y: .Right = X + cw: .Bottom = Y + ch
            End With
            If RectVisible(CDC, Crect) <> 1 Then OutOfScroll = True: Exit Sub
        End If
        
        Dim srcDC As Long
        If Mems(index).GIF.count = 0 Then
            srcDC = Mems(index).hwnd
        Else
            If GetTickCount - Mems(index).GIF.time >= GTick Then Mems(index).GIF.tick = Mems(index).GIF.tick + 1
            If Mems(index).GIF.tick > Mems(index).GIF.count Then Mems(index).GIF.tick = 0
            srcDC = Mems(index).GIF.frames(Mems(index).GIF.tick)
        End If
        
        If IsMissing(alpha) Then
            BitBlt CDC, X, Y, cw, ch, srcDC, cx, cy, vbSrcCopy
        Else
            AlphaBlend CDC, X, Y, cw, ch, srcDC, cx, cy, cw, ch, bl
        End If
        
        With DrawF
            .Left = X
            .top = Y
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
    Public Sub Writes(ByVal text As String, X As Long, Y As Long, Optional size As Long = 14, Optional Color As Long, Optional W As Long = 0, Optional H As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular)
        If OutOfScroll Then Exit Sub
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim ret As Long
            GdipIsVisibleRect gg, X, Y, W, H, ret
            If ret = 0 Then OutOfScroll = True: Exit Sub
        End If
        
        EF.Writes text, X, Y, gg, Color, size, W, H, align, style
        With DrawF
            .Left = X
            .top = Y
            .Right = W
            .Bottom = H
        End With
    End Sub
'========================================================
'   Shape
    Public Sub Paint(ByVal shape As Integer, X As Long, Y As Long, W As Long, H As Long, Optional Color As Long, Optional Radius As Long, Optional style As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        If OutOfScroll Then Exit Sub
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim ret As Long
            GdipIsVisibleRect gg, X - 1, Y - 1, W, H, ret
            If Not ret Then OutOfScroll = True: Exit Sub
        End If
        
        GdipResetPath path
        If shape = 0 Then GdipAddPathRectangle path, X - 1, Y - 1, W, H
        If shape = 1 Then GdipAddPathEllipse path, X, Y, W - 1, H - 1
        
        If Color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then GdipSetPenColor Pen, Color
        End If
        
        If style = 0 Then GdipFillPath gg, brush, path
        If style = 1 Then GdipDrawPath gg, Pen, path
        
        With DrawF
            .Left = X
            .top = Y
            .Right = W
            .Bottom = H
        End With
    End Sub
'========================================================
'   Scroll
    Public Sub StartScroll(X As Long, Y As Long, W As Long, H As Long, SX As Long, SY As Long)
        OutOfScroll = False
        ScrollMode = True
        ScrollX = SX: ScrollY = SY: ScrollBX = X: ScrollBY = Y
        GdipSetClipRect gg, X, Y, W, H, CombineModeReplace
        CRgn = CreateRectRgn(X, Y, X + W, Y + H)
        SelectClipRgn CDC, CRgn
    End Sub
    Public Sub EndScroll()
        OutOfScroll = False
        ScrollMode = False
        GdipResetClip gg
        DeleteObject CRgn
        SelectClipRgn CDC, ByVal 0
    End Sub
'========================================================

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
'Emerald Ïà¹Ø´úÂë

Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
Private Type AniTask
    Start As Long
    during As Long
    custom As Boolean
    func As String
    profunc As Integer
    delay As Long
    mark As Boolean
End Type
Dim GTick As Long
Dim Anis() As AniTask
Dim DrawBox As Object, brush As Long, path As Long, Pen As Long
Dim ScrollMode As Boolean, ScrollX As Long, ScrollY As Long, ScrollBX As Long, ScrollBY As Long, CRgn As Long
Public CDC As Long, GG As Long, OutOfScroll As Boolean
Public Res As GResource
Public TopPage As Boolean
'========================================================
'   Init
    Public Sub Create(dBox As Object)
        ReDim Anis(0)
        Set Res = New GResource
        
        GTick = 1000 / 30
        CDC = CreateCDC(GW, GH)
        Res.Add CDC, 0, ""
        GdipCreateFromHDC CDC, GG
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        
        Set DrawBox = dBox
        
        GdipCreateSolidFill 0, brush
        GdipCreatePen1 0, 1, UnitPixel, Pen
        GdipCreatePath FillModeWinding, path
    End Sub
    Public Sub Dispose()
        GdipDeleteBrush brush
        GdipDeletePen Pen
        GdipDeletePath path
        Res.Dispose
    End Sub
'========================================================
'   Print
    Public Sub Clear(Optional Color As Long = -1)
        GdipGraphicsClear GG, Color
    End Sub
    Public Sub Update()
        Dim i As Integer
        For i = 1 To UBound(Anis)
            If Not Anis(i).mark Then
                If (GetTickCount - Anis(i).Start - Anis(i).delay) > Anis(i).during Then
                     CallByName DrawBox, "AnimationDone", VbMethod, i
                    Anis(i).mark = True
                End If
            End If
        Next
        
        CallByName DrawBox, "Update", VbMethod
    End Sub
    Public Sub Display(DC As Long)
        If TopPage Then
            PaintDC CDC, DC, Alpha:=1
        Else
            BitBlt DC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
        End If
    End Sub
'========================================================
'   RunTime
    Public Property Get SrcClass() As Object
        Set SrcClass = DrawBox
    End Property
    Public Property Get GIFTick() As Long
        GIFTick = GTick
    End Property
    Public Property Let GIFTick(nGIFTick As Long)
        GTick = nGIFTick
    End Property
'========================================================
'   Images
    Public Sub DrawImageEx(n, x As Long, y As Long, Optional w, Optional h, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0)
        Dim index As Integer
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, w, h, 0, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, w, h, 0, pro
                End If

            End If
        End If
        
        If TypeName(n) = "String" Then
            index = Res.GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(w) Then GdipGetImageWidth Res.ResourceImageHandle(index), w
        If IsMissing(h) Then GdipGetImageWidth Res.ResourceImageHandle(index), h
        
        If pos = 1 Then x = Int(x - w / 2): y = Int(y - h / 2)
        If pos = 2 Then x = x - w
        If pos = 3 Then y = y - h
        If pos = 4 Then x = x + w
        If pos = 5 Then y = y + h
        
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, x - 1, y - 1, w, h, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        GdipDrawImageRect GG, Res.ResourceImageHandle(index), x, y, w, h
        
        With DrawF
            .Left = x
            .top = y
            .Right = cw
            .Bottom = ch
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 255, 0, 255)
            GdipDrawRectangle GG, Pen, x, y, w - 1, h - 1
        End If
    End Sub
    Public Sub DrawImage(n, x As Long, y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional Alpha, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        
        If Not IsMissing(Alpha) Then
            Dim pro As Single
            pro = 1
            If animation <> 0 Then
                If Not Anis(animation).mark Then
                    pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                End If
            End If
            Alpha = Alpha * pro
            
            If Alpha < 0 Then Alpha = 0
            If Alpha > 1 Then Alpha = 1
            With b
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(Alpha * 255)
            End With
            CopyMemory bl, b, 4
        End If
        
        If TypeName(n) = "String" Then
            index = Res.GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(cx) Then cx = 0
        If IsMissing(cy) Then cy = 0
        If IsMissing(cw) Then cw = Res.ImgSize(index, imgGetWidth) - cx
        If IsMissing(ch) Then ch = Res.ImgSize(index, imgGetHeight) - cy
        
        If pos = 1 Then x = Int(x - cw / 2): y = Int(y - ch / 2)
        If pos = 2 Then x = x - cw
        If pos = 3 Then y = y - ch
        If pos = 4 Then x = x + cw
        If pos = 5 Then y = y + ch
        
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long, Crect As RECT
            With Crect
                .Left = x: .top = y: .Right = x + cw: .Bottom = y + ch
            End With
            If RectVisible(CDC, Crect) <> 1 Then OutOfScroll = True: Exit Sub
        End If
        
        Dim srcDC As Long
        If Res.ResourceGIFCount(index) = 0 Then
            srcDC = Res.ResourceHandle(index)
        Else
            If GetTickCount - Res.ResourceGIFTime(index) >= GTick Then Res.GIFTick index
            srcDC = Res.ResourceGIFCurrentHandle(index)
        End If
        
        If IsMissing(Alpha) Then
            BitBlt CDC, x, y, cw, ch, srcDC, cx, cy, vbSrcCopy
        Else
            AlphaBlend CDC, x, y, cw, ch, srcDC, cx, cy, cw, ch, bl
        End If
        
        With DrawF
            .Left = x
            .top = y
            .Right = cw
            .Bottom = ch
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 0, 255)
            GdipDrawRectangle GG, Pen, x, y, cw - 1, ch - 1
        End If
    End Sub
'========================================================
'   Writer
    Public Sub Writes(ByVal text As String, x As Long, y As Long, Optional size As Long = 14, Optional Color As Long, Optional w As Long = 0, Optional h As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, x, y, w, h, Ret
            If Ret = 0 Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        EF.Writes text, x, y, GG, Color, size, w, h, align, style
        With DrawF
            .Left = x
            .top = y
            .Right = w
            .Bottom = h
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 255, 0)
            GdipDrawRectangle GG, Pen, x, y, w - 1, h - 1
        End If
    End Sub
'========================================================
'   Shape
    Public Sub PaintLine(X1 As Long, Y1 As Long, x2 As Long, y2 As Long, Optional Color As Long, Optional size As Long = 1, Optional animation As Integer = 0)
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, IIf(X1 < x2, X1, x2), IIf(Y1 < y2, Y1, y2), IIf(X1 < x2, x2 - X1, X1 - x2), IIf(Y1 < y2, y2 - Y1, Y1 - y2), Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        If Color <> 0 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, size
        End If
        
        GdipDrawLine GG, Pen, X1, Y1, x2, y2
        
        With DrawF
            .Left = IIf(X1 < x2, X1, x2)
            .top = IIf(Y1 < y2, Y1, y2)
            .Right = IIf(X1 < x2, x2 - X1, X1 - x2)
            .Bottom = IIf(Y1 < y2, y2 - Y1, Y1 - y2)
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 240, 176, 0)
            GdipDrawRectangle GG, Pen, DrawF.Left, DrawF.top, DrawF.Right + 1, DrawF.Bottom + 1
        End If
    End Sub
    Public Sub Paint(ByVal shape As Integer, x As Long, y As Long, w As Long, h As Long, Optional Color As Long, Optional Radius As Long, Optional size As Long = 1, Optional style As Integer = 0, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        'If OutOfScroll Then Exit Sub
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, x - 1, y - 1, w, h, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath path
        
        If pos = 1 Then x = Int(x - w / 2): y = Int(y - h / 2)
        If pos = 2 Then x = x - w
        If pos = 3 Then y = y - h
        If pos = 4 Then x = x + w
        If pos = 5 Then y = y + h
        
ReShape:
        If shape = 0 Then GdipAddPathRectangle path, x, y, w - 1, h - 1
        If shape = 1 Then GdipAddPathEllipse path, x, y, w - 1, h - 1
        If shape = 2 Then
        
            If Radius = 0 Then
                shape = 0: GoTo ReShape
            End If
            
            If Radius > w Then Radius = w
            If Radius > h Then Radius = h
            
            GdipAddPathArc path, x, y, Radius, Radius, 180, 90
            GdipAddPathArc path, x + w - Radius, y, Radius, Radius, 270, 90
            GdipAddPathArc path, x + w - Radius, y + h - Radius, Radius, Radius, 0, 90
            GdipAddPathArc path, x, y + h - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure path
            
        End If
        
        If Color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, size
            End If
        End If
        
        If shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeDefault
        If style = 0 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        If shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        
        With DrawF
            .Left = x
            .top = y
            .Right = w
            .Bottom = h
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 255, 0, 0)
            GdipDrawRectangle GG, Pen, x, y, w - 1, h - 1
        End If
    End Sub
    Public Sub PaintArc(x As Long, y As Long, w As Long, h As Long, degree As Long, Optional Start As Long = 0, Optional Color As Long, Optional size As Long = 1, Optional style As Integer = 0, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'style:0=fill,1=border,2=sector
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, x - 1, y - 1, w, h, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, w, h, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath path
        
        If pos = 1 Then x = Int(x - w / 2): y = Int(y - h / 2)
        If pos = 2 Then x = x - w
        If pos = 3 Then y = y - h
        If pos = 4 Then x = x + w
        If pos = 5 Then y = y + h
        
        GdipAddPathArc path, x, y, w, h, Start, degree
        If style = 2 Then GdipAddPathLine path, x + w / 2, y + w / 2, x + w / 2, y + w / 2
        If style = 0 Then GdipClosePathFigure path
        
        If Color <> 0 Then
            If style = 0 Or style = 2 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, size
            End If
        End If
        
        If style = 0 Or style = 2 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        
        With DrawF
            .Left = x
            .top = y
            .Right = w
            .Bottom = h
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 255, 255)
            GdipDrawRectangle GG, Pen, x, y, w - 1, h - 1
        End If
    End Sub
    Public Sub PaintPolygon(Color As Long, style As Integer, ParamArray Points())

        GdipResetPath path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For i = 0 To UBound(Points) Step 2
            p(i / 2).x = Points(i): p(i / 2).y = Points(i + 1)
        Next
        
        GdipAddPathPolygon path, p(0), UBound(p) + 1
        
        If style = 0 Then GdipSetSolidFillColor brush, Color
        If style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, size
        End If
        
        If style = 0 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, path
        End If
    End Sub
'========================================================
'   Scroll
    Public Sub StartScroll(x As Long, y As Long, w As Long, h As Long, sx As Long, sy As Long)
        OutOfScroll = False
        ScrollMode = True
        ScrollX = sx: ScrollY = sy: ScrollBX = x: ScrollBY = y
        GdipSetClipRect GG, x, y, w, h, CombineModeReplace
        CRgn = CreateRectRgn(x, y, x + w, y + h)
        SelectClipRgn CDC, CRgn
    End Sub
    Public Sub EndScroll()
        OutOfScroll = False
        ScrollMode = False
        GdipResetClip GG
        DeleteObject CRgn
        SelectClipRgn CDC, ByVal 0
    End Sub
'========================================================
    'Animation
    Public Function NewAnimation(custom As Boolean, func As String, profunc As Integer, during As Long) As Integer
        ReDim Preserve Anis(UBound(Anis) + 1)
        With Anis(UBound(Anis))
            .custom = custom
            .during = during
            .func = func
            .profunc = profunc
            .mark = True
        End With
        
        NewAnimation = UBound(Anis)
    End Function
    Public Sub StartAnimation(id As Integer, Optional delay As Long = 0)
        Anis(id).Start = GetTickCount
        Anis(id).delay = delay
        Anis(id).mark = False
    End Sub
    Public Function AnimationDone(id As Integer) As Boolean
        AnimationDone = Anis(id).mark
    End Function
'========================================================
    'Control
    Public Sub ShowLoading(x As Long, y As Long, w As Long, h As Long, size As Long, color1 As Long, color2 As Long, color3 As Long)
        
        PaintArc x, y, w, h, 120, (GetTickCount Mod 3000) / 3000 * 360, color3, size * 0.86, 1
        PaintArc x, y, w, h, 120, (GetTickCount Mod 1800) / 1800 * 360 + 90, color2, size * 0.93, 1
        PaintArc x, y, w, h, 120, (GetTickCount Mod 1200) / 1200 * 360 + 90, color1, size, 1
        'argb(255, 242, 195, 17),argb(255, 240, 127, 34),argb(255, 232, 76, 61)
    End Sub
    Public Function ShowButton(pic As String, x As Long, y As Long, text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim w As Long, h As Long, i As Integer, m As Integer
        i = Res.GetImage(pic)
        w = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(x, y, w, Int(h / 2))
        DrawImage i, x, y, cy:=IIf(m <> 0, Int(h / 2), 0), ch:=Int(h / 2), Alpha:=1
        
        Writes text, x, y + h / 2 / 2 - size / 0.75 / 2, size, Color, w, size / 0.75, StringAlignmentCenter, style
        
        ShowButton = m
    End Function
    Public Function ShowSimpleButton(pic As String, x As Long, y As Long, text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim w As Long, h As Long, i As Integer, m As Integer
        i = Res.GetImage(pic)
        w = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(x, y, w, h)
        DrawImage i, x, y, Alpha:=IIf(m, 1, 0.8)
        
        Writes text, x, y + h / 2 - size / 0.75 / 2, size, Color, w, size / 0.75, StringAlignmentCenter, style
    
        ShowSimpleButton = m
    End Function
    Public Function ShowColorButton(shape As Integer, x As Long, y As Long, w As Long, h As Long, text As String, TextColor As Long, Color As Long, HoverColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim m As Integer
        
        m = CheckMouse(x, y, w, h)
        Paint shape, x, y, w, h, IIf(m, HoverColor, Color), Radius, style:=ShapeStyle
        
        Writes text, x, y + h / 2 - size / 0.75 / 2, size, TextColor, w, size / 0.75, StringAlignmentCenter, style
    
        ShowColorButton = m
    End Function
    Public Function ShowCheckBox(value, pic As String, x As Long, y As Long, text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim w As Long, h As Long, i As Integer, m As Integer
        Dim x2 As Long, y2 As Long
        x2 = x: y2 = y
        
        i = Res.GetImage(pic)
        w = Res.ImgSize(i, imgGetWidth): h = Res.ImgSize(i, imgGetHeight)
        
        m = CheckMouse(x, y, w, Int(h / 2))
        DrawImage i, x, y, cy:=IIf(value = True, Int(h / 2), 0), ch:=Int(h / 2), Alpha:=1
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes text, x2 + w + size / 0.75 / 2, y2 + h / 2 / 2 - size / 0.75 / 2, size, Color, GW, size / 0.75, StringAlignmentNear, style
        
        ShowCheckBox = m
    End Function
    Public Function ShowHScrollBar(value, x As Long, y As Long, w As Long, h As Long, LineColor As Long, BallColor As Long, Optional size As Long = 2) As Integer
        Dim m As Integer
        
        m = CheckMouse(x, y, w, h)
        PaintLine x, y + h / 2 - size / 2, x + w, y + h / 2 - size / 2, LineColor, size
        Paint 1, x + value * (w - h), y, h, h, BallColor
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.x - x) / (w - h)
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowHScrollBar = m
    End Function
    Public Function ShowVScrollBar(value, x As Long, y As Long, w As Long, h As Long, LineColor As Long, BallColor As Long, Optional size As Long = 2) As Integer
        Dim m As Integer
        
        m = CheckMouse(x, y, w, h)
        PaintLine x + w / 2 - size / 2, y, x + w / 2 - size / 2, y + h, LineColor, size
        Paint 1, x, y + value * (h - w), w, w, BallColor
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.y - y) / (h - w)
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowVScrollBar = m
    End Function
'========================================================

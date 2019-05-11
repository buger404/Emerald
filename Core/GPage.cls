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
Dim DrawBox As Object, brush As Long, Path As Long, Pen As Long
Dim ScrollMode As Boolean, ScrollX As Long, ScrollY As Long, ScrollBX As Long, ScrollBY As Long, CRgn As Long
Public CDC As Long, GG As Long, OutOfScroll As Boolean
Public Res As GResource
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
        GdipCreatePath FillModeWinding, Path
    End Sub
    Public Sub Dispose()
        GdipDeleteBrush brush
        GdipDeletePen Pen
        GdipDeletePath Path
        Res.Dispose
    End Sub
'========================================================
'   Print
    Public Sub Clear()
        GdipGraphicsClear GG, argb(255, 255, 255, 255)
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
        BitBlt DC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
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
    Public Sub DrawImageEx(n, X As Long, Y As Long, Optional W, Optional H, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
        Dim index As Integer
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, W, H, 0, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, W, H, 0, pro
                End If

            End If
        End If
        
        If TypeName(n) = "String" Then
            index = Res.GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(W) Then GdipGetImageWidth Res.ResourceImageHandle(index), W
        If IsMissing(H) Then GdipGetImageWidth Res.ResourceImageHandle(index), H
        
        If Pos = 1 Then X = Int(X - W / 2): Y = Int(Y - H / 2)
        If Pos = 2 Then X = X - W
        If Pos = 3 Then Y = Y - H
        If Pos = 4 Then X = X + W
        If Pos = 5 Then Y = Y + H
        
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, X - 1, Y - 1, W, H, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        GdipDrawImageRect GG, Res.ResourceImageHandle(index), X, Y, W, H
        
        With DrawF
            .Left = X
            .top = Y
            .Right = cw
            .Bottom = ch
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 255, 0, 255)
            GdipDrawRectangle GG, Pen, X, Y, W - 1, H - 1
        End If
    End Sub
    Public Sub DrawImage(n, X As Long, Y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional Alpha, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        
        If animation <> 0 Then
            If (GetTickCount - Anis(animation).Start) > Anis(animation).during Then animation = 0
        End If
        
        If Not IsMissing(Alpha) Then
            If animation = 0 Then
setAlpha:
                If Alpha < 0 Then Alpha = 0
                If Alpha > 1 Then Alpha = 1
                With b
                    .AlphaFormat = &H1
                    .BlendFlags = &H0
                    .BlendOp = 0
                    .SourceConstantAlpha = Int(Alpha * 255)
                End With
                CopyMemory bl, b, 4
                If animation <> 0 Then Return
            End If
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
        
        If Pos = 1 Then X = Int(X - cw / 2): Y = Int(Y - ch / 2)
        If Pos = 2 Then X = X - cw
        If Pos = 3 Then Y = Y - ch
        If Pos = 4 Then X = X + cw
        If Pos = 5 Then Y = Y + ch
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, ((GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during))
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, cw, ch, Alpha, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, cw, ch, Alpha, pro
                End If
                
                GoSub setAlpha
            End If
        End If
        
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long, Crect As RECT
            With Crect
                .Left = X: .top = Y: .Right = X + cw: .Bottom = Y + ch
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
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 0, 255)
            GdipDrawRectangle GG, Pen, X, Y, cw - 1, ch - 1
        End If
    End Sub
'========================================================
'   Writer
    Public Sub Writes(ByVal text As String, X As Long, Y As Long, Optional size As Long = 14, Optional Color As Long, Optional W As Long = 0, Optional H As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, X, Y, W, H, Ret
            If Ret = 0 Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        EF.Writes text, X, Y, GG, Color, size, W, H, align, style
        With DrawF
            .Left = X
            .top = Y
            .Right = W
            .Bottom = H
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 255, 0)
            GdipDrawRectangle GG, Pen, X, Y, W - 1, H - 1
        End If
    End Sub
'========================================================
'   Shape
    Public Sub PaintLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Optional Color As Long, Optional size As Long = 1, Optional animation As Integer = 0)
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, IIf(X1 < X2, X1, X2), IIf(Y1 < Y2, Y1, Y2), IIf(X1 < X2, X2 - X1, X1 - X2), IIf(Y1 < Y2, Y2 - Y1, Y1 - Y2), Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        If Color <> 0 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, size
        End If
        
        GdipDrawLine GG, Pen, X1, Y1, X2, Y2
        
        With DrawF
            .Left = IIf(X1 < X2, X1, X2)
            .top = IIf(Y1 < Y2, Y1, Y2)
            .Right = IIf(X1 < X2, X2 - X1, X1 - X2)
            .Bottom = IIf(Y1 < Y2, Y2 - Y1, Y1 - Y2)
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 240, 176, 0)
            GdipDrawRectangle GG, Pen, DrawF.Left, DrawF.top, DrawF.Right + 1, DrawF.Bottom + 1
        End If
    End Sub
    Public Sub Paint(ByVal shape As Integer, X As Long, Y As Long, W As Long, H As Long, Optional Color As Long, Optional Radius As Long, Optional size As Long = 1, Optional style As Integer = 0, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        'If OutOfScroll Then Exit Sub
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, X - 1, Y - 1, W, H, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath Path
        
        If Pos = 1 Then X = Int(X - W / 2): Y = Int(Y - H / 2)
        If Pos = 2 Then X = X - W
        If Pos = 3 Then Y = Y - H
        If Pos = 4 Then X = X + W
        If Pos = 5 Then Y = Y + H
        
ReShape:
        If shape = 0 Then GdipAddPathRectangle Path, X, Y, W - 1, H - 1
        If shape = 1 Then GdipAddPathEllipse Path, X, Y, W - 1, H - 1
        If shape = 2 Then
        
            If Radius = 0 Then
                shape = 0: GoTo ReShape
            End If
            
            If Radius > W Then Radius = W
            If Radius > H Then Radius = H
            
            GdipAddPathArc Path, X, Y, Radius, Radius, 180, 90
            GdipAddPathArc Path, X + W - Radius, Y, Radius, Radius, 270, 90
            GdipAddPathArc Path, X + W - Radius, Y + H - Radius, Radius, Radius, 0, 90
            GdipAddPathArc Path, X, Y + H - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure Path
            
        End If
        
        If Color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, size
            End If
        End If
        
        If shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeDefault
        If style = 0 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        If shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        
        With DrawF
            .Left = X
            .top = Y
            .Right = W
            .Bottom = H
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 255, 0, 0)
            GdipDrawRectangle GG, Pen, X, Y, W - 1, H - 1
        End If
    End Sub
    Public Sub PaintArc(X As Long, Y As Long, W As Long, H As Long, Degree As Long, Optional Start As Long = 0, Optional Color As Long, Optional size As Long = 1, Optional style As Integer = 0, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'style:0=fill,1=border,2=sector
        If ScrollMode Then
            X = X + ScrollX + ScrollBX: Y = Y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, X - 1, Y - 1, W, H, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, X, Y, W, H, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath Path
        
        If Pos = 1 Then X = Int(X - W / 2): Y = Int(Y - H / 2)
        If Pos = 2 Then X = X - W
        If Pos = 3 Then Y = Y - H
        If Pos = 4 Then X = X + W
        If Pos = 5 Then Y = Y + H
        
        GdipAddPathArc Path, X, Y, W, H, Start, Degree
        If style = 2 Then GdipAddPathLine Path, X + W / 2, Y + W / 2, X + W / 2, Y + W / 2
        If style = 0 Then GdipClosePathFigure Path
        
        If Color <> 0 Then
            If style = 0 Or style = 2 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, size
            End If
        End If
        
        If style = 0 Or style = 2 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        
        With DrawF
            .Left = X
            .top = Y
            .Right = W
            .Bottom = H
        End With
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 255, 255)
            GdipDrawRectangle GG, Pen, X, Y, W - 1, H - 1
        End If
    End Sub
    Public Sub PaintPolygon(Color As Long, style As Integer, ParamArray Points())

        GdipResetPath Path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For i = 0 To UBound(Points) Step 2
            p(i / 2).X = Points(i): p(i / 2).Y = Points(i + 1)
        Next
        
        GdipAddPathPolygon Path, p(0), UBound(p) + 1
        
        If style = 0 Then GdipSetSolidFillColor brush, Color
        If style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, size
        End If
        
        If style = 0 Then GdipFillPath GG, brush, Path
        If style = 1 Then GdipDrawPath GG, Pen, Path
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, Path
        End If
    End Sub
'========================================================
'   Scroll
    Public Sub StartScroll(X As Long, Y As Long, W As Long, H As Long, sx As Long, sy As Long)
        OutOfScroll = False
        ScrollMode = True
        ScrollX = sx: ScrollY = sy: ScrollBX = X: ScrollBY = Y
        GdipSetClipRect GG, X, Y, W, H, CombineModeReplace
        CRgn = CreateRectRgn(X, Y, X + W, Y + H)
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
    Public Function ShowButton(pic As String, X As Long, Y As Long, text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim W As Long, H As Long, i As Integer, m As Integer
        i = GetImage(pic)
        W = Mems(i).W: H = Mems(i).H
        
        m = CheckMouse(X, Y, W, Int(H / 2))
        DrawImage i, X, Y, cy:=IIf(m <> 0, Int(H / 2), 0), ch:=Int(H / 2), Alpha:=1
        
        Writes text, X, Y + H / 2 / 2 - size / 0.75 / 2, size, Color, W, size / 0.75, StringAlignmentCenter, style
        
        ShowButton = m
    End Function
    Public Function ShowSimpleButton(pic As String, X As Long, Y As Long, text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim W As Long, H As Long, i As Integer, m As Integer
        i = GetImage(pic)
        W = Mems(i).W: H = Mems(i).H
        
        m = CheckMouse(X, Y, W, H)
        DrawImage i, X, Y, Alpha:=IIf(m, 1, 0.8)
        
        Writes text, X, Y + H / 2 - size / 0.75 / 2, size, Color, W, size / 0.75, StringAlignmentCenter, style
    
        ShowSimpleButton = m
    End Function
    Public Function ShowColorButton(shape As Integer, X As Long, Y As Long, W As Long, H As Long, text As String, TextColor As Long, Color As Long, HoverColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim m As Integer
        
        m = CheckMouse(X, Y, W, H)
        Paint shape, X, Y, W, H, IIf(m, HoverColor, Color), Radius, style:=ShapeStyle
        
        Writes text, X, Y + H / 2 - size / 0.75 / 2, size, TextColor, W, size / 0.75, StringAlignmentCenter, style
    
        ShowColorButton = m
    End Function
    Public Function ShowCheckBox(value, pic As String, X As Long, Y As Long, text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As Integer
        Dim W As Long, H As Long, i As Integer, m As Integer
        i = GetImage(pic)
        W = Mems(i).W: H = Mems(i).H
        
        m = CheckMouse(X, Y, W, Int(H / 2))
        DrawImage i, X, Y, cy:=IIf(value = True, Int(H / 2), 0), ch:=Int(H / 2), Alpha:=1
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes text, X + W + size / 0.75 / 2, Y + H / 2 / 2 - size / 0.75 / 2, size, Color, GW, size / 0.75, StringAlignmentNear, style
        
        ShowCheckBox = m
    End Function
    Public Function ShowHScrollBar(value, X As Long, Y As Long, W As Long, H As Long, LineColor As Long, BallColor As Long, Optional size As Long = 2) As Integer
        Dim m As Integer
        
        m = CheckMouse(X, Y, W, H)
        PaintLine X, Y + H / 2 - size / 2, X + W, Y + H / 2 - size / 2, LineColor, size
        Paint 1, X + value * (W - H), Y, H, H, BallColor
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.X - X) / (W - H)
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowHScrollBar = m
    End Function
'========================================================

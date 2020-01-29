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
'Emerald 相关代码

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
Dim CRgn As Long
Dim AniCollection() As EAnimation, PlayAni() As EAnimation
Public OODC As Long, OOGG As Long, ScrollMode As Boolean, ScrollWidth As Long, ScrollHeight As Long
Public CDC As Long, GG As Long
Public Res As GResource
Public TopPage As Boolean
Public IsSystem As Boolean
'========================================================
'   Init
    Public Sub Create(dBox As Object)
        ReDim Anis(0)
        ReDim AniCollection(0): ReDim PlayAni(0)
        Set Res = New GResource
        
        GTick = 1000 / 30
        CDC = CreateCDC(RGW, RGH)
        Res.Add CDC, 0, ""
        PoolCreateFromHdc CDC, GG
        
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        
        Set DrawBox = dBox
        
        PoolCreateSolidFill 0, brush
        PoolCreatePen1 0, 1, UnitPixel, Pen
        PoolCreatePath FillModeWinding, path
    End Sub
    Public Sub Dispose()
        Exit Sub
        PoolDeleteBrush brush
        PoolDeletePen Pen
        GdipDeletePath path
        Res.Dispose
    End Sub
'========================================================
'   Print
    Public Sub Clear(Optional Color As Long = -1)
        GdipGraphicsClear GG, Color
    End Sub
    Public Sub Update()
        Dim I As Integer
        For I = 1 To UBound(Anis)
            If Not Anis(I).mark Then
                If (GetTickCount - Anis(I).Start - Anis(I).delay) > Anis(I).during Then
                    Call DrawBox.AnimationDone(I)
                    Anis(I).mark = True
                End If
            End If
        Next
        
        'On Error Resume Next
        Call DrawBox.Update
        'If Err.Number <> 0 Then Call ECore.RaiseErrorScreen: Exit Sub
    End Sub
    Public Sub Display(DC As Long)
        If TopPage Then
            If Not ECore.FreezeAvailable Then PaintDC CDC, DC, Alpha:=1
        Else
            If Not ECore.FreezeAvailable Then BitBlt DC, 0, 0, RGW, RGH, CDC, 0, 0, vbSrcCopy
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
    Public Sub LoadAnimation(path As String)
        ReDim Preserve AniCollection(UBound(AniCollection) + 1)
        AniCollection(UBound(AniCollection)) = Animations.LoadAnimation(path)
        With AniCollection(UBound(AniCollection))
            For I = 1 To UBound(.channel)
                For S = 1 To UBound(.channel(I).frames)
                    For t = 0 To UBound(.channel(I).frames(S).aframes)
                        .channel(I).frames(S).aframes(t).picindex = Res.GetImage(.channel(I).frames(S).aframes(t).pic)
                    Next
                Next
            Next
        End With
    End Sub
    Public Sub LoadAnimationsFromDir(path As String)
        Dim f As String
        If Right(path, 1) <> "\" Then path = path & "\"
        
        f = Dir(path & "*.emrani")
        Do While f <> ""
            Call LoadAnimation(path & f)
            f = Dir()
        Loop
    End Sub
    Private Function FindAnimation(Name As String) As EAnimation
        For I = 1 To UBound(AniCollection)
            If AniCollection(I).Name = Name Then FindAnimation = AniCollection(I): Exit For
        Next
    End Function
    Private Function GetPlayAnimation(playname As String) As Integer
        For I = 1 To UBound(PlayAni)
            If PlayAni(I).Name = playname Then GetPlayAnimation = I: Exit For
        Next
    End Function
    Public Sub CreatePlayAnimation(Name As String, playname As String, channel As String)
        ReDim Preserve PlayAni(UBound(PlayAni) + 1)
        PlayAni(UBound(PlayAni)) = FindAnimation(Name)
        PlayAni(UBound(PlayAni)).Name = playname
        PlayAni(UBound(PlayAni)).tick = GetTickCount
        For I = 1 To UBound(PlayAni(UBound(PlayAni)).channel)
            PlayAni(UBound(PlayAni)).channel(I).CurrentFrame = 1
            If PlayAni(UBound(PlayAni)).channel(I).Name = channel Then PlayAni(UBound(PlayAni)).CurrentChannel = I
        Next
    End Sub
    Public Sub DisposePlayAnimation(playname As String)
        PlayAni(GetPlayAnimation(playname)) = PlayAni(UBound(PlayAni))
        ReDim Preserve PlayAni(UBound(PlayAni) - 1)
    End Sub
    Public Sub SwitchChannel(playname As String, channel As String)
        Dim S As Integer
        S = GetPlayAnimation(playname)
        If S = 0 Then Exit Sub
        If PlayAni(S).channel(PlayAni(S).CurrentChannel).frames(PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame).tick = -1 Then
            PlayAni(S).channel(PlayAni(S).CurrentChannel).frames(PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame).tick = 0
        End If
        
        For I = 1 To UBound(PlayAni(S).channel)
            If PlayAni(S).channel(I).Name = channel Then
                PlayAni(S).channel(I).CurrentFrame = 1
                PlayAni(S).CurrentChannel = I
                PlayAni(S).tick = GetTickCount
                Exit For
            End If
        Next
    End Sub
'========================================================
'   Images
    Public Sub DrawAnimation(playname As String, X As Long, Y As Long, Optional Direction As ImgDirection = DirNormal)
        If ECore.FreezeAvailable Then Exit Sub
        
        Dim S As Integer, m As Integer, W As Long, H As Long, last As Boolean
        
        S = GetPlayAnimation(playname)
        If S = 0 Then Exit Sub
        
        With PlayAni(S).channel(PlayAni(S).CurrentChannel).frames(PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame)
            For I = 0 To UBound(.aframes)
                If .aframes(I).Alpha = 1 Then m = 1
                If .aframes(I).Alpha <> 1 Then m = 1
                If .aframes(I).size <> 1 Then m = 0
                
                If m = 0 Then
                    W = Res.ImgSize(.aframes(I).picindex, imgGetWidth) * .aframes(I).size
                    H = Res.ImgSize(.aframes(I).picindex, imgGetHeight) * .aframes(I).size
                    DrawImageEx .aframes(I).picindex, X + .aframes(I).X, Y + .aframes(I).Y, W, H, PlayAni(S).position
                Else
                    DrawImage .aframes(I).picindex, X + .aframes(I).X, Y + .aframes(I).Y, pos:=PlayAni(S).position, Alpha:=.aframes(I).Alpha, Direction:=Direction
                End If
            Next
            
            Dim DisposeMark As Boolean
            If .tick <> -1 Then
                last = (PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame = UBound(PlayAni(S).channel(PlayAni(S).CurrentChannel).frames))
                If GetTickCount - PlayAni(S).tick >= .tick Then
                    If .msg = "dispose" Or .disposed Then DisposeMark = True
                    If .msg <> "" Then
                        CallByName SrcClass, "AnimationMsg", VbMethod, playname, .msg
                    End If
                    If .tick <> 0 Then
                        If last Then
                            PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame = 1
                        Else
                            PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame = PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame + 1
                        End If
                    End If
                    If .tick = 0 Then .tick = -1
                    PlayAni(S).tick = GetTickCount
                    If Not (EMusic Is Nothing) Then
                        For I = 0 To UBound(PlayAni(S).channel(PlayAni(S).CurrentChannel).frames(PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame).sounds)
                            EMusic.Play PlayAni(S).channel(PlayAni(S).CurrentChannel).frames(PlayAni(S).channel(PlayAni(S).CurrentChannel).CurrentFrame).sounds(I).snd
                        Next
                    End If
                End If
            End If
        End With
        
        If DisposeMark Then DisposePlayAnimation playname
    End Sub
    Public Sub DrawImageEx(n, ByVal X As Single, ByVal Y As Single, Optional W, Optional H, Optional pos As PosAlign = posNormal, Optional angle As Long = 0, Optional animation As Integer = 0)
        Dim index As Integer
        
        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub
        
        If Scales <> 1 Then
            X = X * Scales: Y = Y * Scales
            If Not IsMissing(W) Then W = W * Scales
            If Not IsMissing(H) Then H = H * Scales
        End If
        
        Dim OX As Single, OY As Single, ow As Long, oh As Long
        OX = X: OY = Y: ow = W: oh = H
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, ow, oh, 0, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, ow, oh, 0, pro
                End If

            End If
        End If
        
        If TypeName(n) = "String" Then
            index = Res.GetImage(n)
        Else
            index = n
        End If
         
        If IsMissing(ow) Then GdipGetImageWidth Res.ResourceImageHandle(index), ow
        If IsMissing(oh) Then GdipGetImageWidth Res.ResourceImageHandle(index), oh
        
        If pos = 1 Then OX = Int(OX - ow / 2): OY = Int(OY - oh / 2)
        If pos = 2 Then OX = OX - ow
        If pos = 3 Then OY = OY - oh
        If pos = 4 Then OX = OX + ow
        If pos = 5 Then OY = OY + oh
        
        If index = 0 Then
            Paint 0, OX, OY, ow, oh, argb(255, 255, 0, 0), size:=2, style:=1
            PaintLine OX, OY, OX + ow, OY + oh, argb(170, 255, 0, 0), 2
            PaintLine OX + ow, OY, OX, OY + oh, argb(170, 255, 0, 0), 2
            Writes n & " 无法找到", OX + ow / 2, OY + oh / 2, 18, argb(255, 255, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
            Exit Sub
        End If
        
        If angle <> 0 Then
            GdipRotateWorldTransform GG, angle, MatrixOrderAppend
            GdipTranslateWorldTransform GG, OX, OY, MatrixOrderAppend
            GdipDrawImageRect GG, Res.ResourceImageHandle(index), -ow / 2, -oh / 2, ow, oh
            GdipResetWorldTransform GG
        Else
            GdipDrawImageRect GG, Res.ResourceImageHandle(index), OX, OY, ow, oh
        End If
        
        With DrawF
            .X = OX
            .Y = OY
            .Width = ow
            .Height = oh
            .CrashIndex = Res.ResourceCrashIndex(index)
            .Shape = -1
            .WSc = Res.ImgSize(index, imgGetWidth) / .Width
            .HSc = Res.ImgSize(index, imgGetHeight) / .Height
        End With
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 255, 0, 255)
            GdipDrawRectangle GG, Pen, OX, OY, ow - 1, oh - 1
        End If
        
        If Debug_pos Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
        
        If ChoosePosition Then
            Call JudgeChoosePosition(OX, OY, ow, oh)
        End If
    End Sub
    Public Sub DrawImage(n, ByVal X As Long, ByVal Y As Long, Optional CX, Optional CY, Optional cw, Optional ch, Optional Alpha, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0, Optional Direction As ImgDirection = DirNormal)
        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub
        
        'If OutOfScroll Then Exit Sub
        If Scales <> 1 Then
            X = X * Scales: Y = Y * Scales
            If Not IsMissing(CX) Then CX = CX * Scales
            If Not IsMissing(CY) Then CY = CY * Scales
            If Not IsMissing(cw) Then cw = cw * Scales
            If Not IsMissing(ch) Then ch = ch * Scales
        End If
        Dim B As BLENDFUNCTION, index As Integer, bl As Long
        Dim OX As Long, OY As Long, ow As Long, oh As Long
        OX = X: OY = Y: ow = W: oh = H
        Dim ocx As Long, ocy As Long, ocw As Long, och As Long
        
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
            With B
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(Alpha * 255)
            End With
            CopyMemory bl, B, 4
        End If
        
        If TypeName(n) = "String" Then
            index = Res.GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(CX) Then ocx = 0 Else: ocx = CX
        If IsMissing(CY) Then ocy = 0 Else: ocy = CY
        If IsMissing(cw) Then ocw = Res.ImgSize(index, imgGetWidth) - ocx Else: ocw = cw
        If IsMissing(ch) Then och = Res.ImgSize(index, imgGetHeight) - ocy Else: och = ch
        
        If pos = 1 Then OX = Int(OX - ocw / 2): OY = Int(OY - och / 2)
        If pos = 2 Then OX = OX - ocw
        If pos = 3 Then OY = OY - och
        If pos = 4 Then OX = OX + ocw
        If pos = 5 Then OY = OY + och
        
        If index = 0 Then
            Writes n & " 无法找到", OX, OY, 18, argb(255, 255, 0, 0), align:=StringAlignmentCenter, style:=FontStyleBold
            Exit Sub
        End If
        
        Dim srcDC As Long
        If Res.ResourceGIFCount(index) = 0 Then
            srcDC = Res.ResourceHandle2(index, Direction)
        Else
            If GetTickCount - Res.ResourceGIFTime(index) >= GTick Then Res.GIFTick index
            srcDC = Res.ResourceGIFCurrentHandle(index)
        End If
        
        If IsMissing(Alpha) Then
            BitBlt CDC, OX, OY, ocw, och, srcDC, ocx, ocy, vbSrcCopy
        Else
            AlphaBlend CDC, OX, OY, ocw, och, srcDC, ocx, ocy, ocw, och, bl
        End If

        With DrawF
            .X = OX
            .Y = OY
            .Width = ocw
            .Height = och
            .CrashIndex = Res.ResourceCrashIndex(index)
            .WSc = 1
            .HSc = 1
            .Shape = -1
        End With
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 0, 255)
            GdipDrawRectangle GG, Pen, OX, OY, ocw - 1, och - 1
        End If
        
        If Debug_pos Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
        
        If ChoosePosition Then
            Call JudgeChoosePosition(OX, OY, ocw, och)
        End If
    End Sub
'========================================================
'   Writer
    Public Sub Writes(ByVal Text As String, ByVal X As Single, ByVal Y As Single, Optional size As Long = 14, Optional Color As Long, Optional ByVal W As Single = 0, Optional ByVal H As Single = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        
        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub
        
        If Scales <> 1 Then
            X = X * Scales: Y = Y * Scales
            size = size * Scales
            If Not IsMissing(W) Then W = W * Scales
            If Not IsMissing(H) Then H = H * Scales
        End If
        
        Dim OX As Single, OY As Single
        OX = X: OY = Y
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, W, H, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, W, H, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        EF.Writes Text, OX, OY, GG, Color, size, W, H, align, style
        With DrawF
            .X = OX
            .Y = OY
            .Width = W
            .Height = H
            .CrashIndex = 0
            .Shape = 0
        End With
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 255, 0)
            GdipDrawRectangle GG, Pen, OX, OY, W - 1, H - 1
        End If
        
        If Debug_pos Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
        
        If ChoosePosition Then
            Call JudgeChoosePosition(OX, OY, W, H)
        End If
    End Sub
'========================================================
'   Shape
    Public Sub PaintLine(ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, Optional Color As Long, Optional size As Long = 1, Optional animation As Integer = 0)
        
        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub
        
        If Scales <> 1 Then
            X1 = X1 * Scales: Y1 = Y1 * Scales
            X2 = X2 * Scales
            Y2 = Y2 * Scales
            size = size * Scales
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
            .X = IIf(X1 < X2, X1, X2)
            .Y = IIf(Y1 < Y2, Y1, Y2)
            .Width = IIf(X1 < X2, X2 - X1, X1 - X2)
            .Height = IIf(Y1 < Y2, Y2 - Y1, Y1 - Y2)
            .CrashIndex = 0
            .Shape = 0
        End With
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 240, 176, 0)
            GdipDrawRectangle GG, Pen, DrawF.X, DrawF.Y, DrawF.Width + 1, DrawF.Height + 1
        End If
        
        If ChoosePosition Then
            Call JudgeChoosePosition(X1, Y1, 0, 0)
            Call JudgeChoosePosition(X2, Y2, 0, 0)
            Call JudgeChoosePosition(DrawF.X, DrawF.Y, DrawF.Width, DrawF.Height)
        End If
    End Sub
    Public Sub Paint(ByVal Shape As Integer, ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, Optional Color As Long, Optional Radius As Long, Optional size As Long = 1, Optional style As Integer = 0, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        'If OutOfScroll Then Exit Sub

        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub

        If Scales <> 1 Then
            X = X * Scales: Y = Y * Scales
            W = W * Scales
            H = H * Scales
            size = size * Scales
        End If

        Dim OX As Single, OY As Single, ow As Single, oh As Single
        OX = X: OY = Y: ow = W: oh = H
        
        If ScrollMode Then
            OX = OX + ScrollX + ScrollBX: OY = OY + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, OX - 1, OY - 1, ow, oh, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath path
        
        If pos = 1 Then OX = Int(OX - ow / 2): OY = Int(OY - oh / 2)
        If pos = 2 Then OX = OX - ow
        If pos = 3 Then OY = OY - oh
        If pos = 4 Then OX = OX + ow
        If pos = 5 Then OY = OY + oh
        
ReShape:
        If Shape = 0 Then GdipAddPathRectangle path, OX, OY, ow - 1, oh - 1
        If Shape = 1 Then GdipAddPathEllipse path, OX, OY, ow - 1, oh - 1
        If Shape = 2 Then
        
            If Radius = 0 Then
                Shape = 0: GoTo ReShape
            End If
            
            If Radius > ow Then Radius = ow
            If Radius > oh Then Radius = oh
            
            GdipAddPathArc path, OX, OY, Radius, Radius, 180, 90
            GdipAddPathArc path, OX + ow - Radius, OY, Radius, Radius, 270, 90
            GdipAddPathArc path, OX + ow - Radius, OY + oh - Radius, Radius, Radius, 0, 90
            GdipAddPathArc path, OX, OY + oh - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure path
            
        End If
        
        If Color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, size
            End If
        End If
        
        If Shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeDefault
        If style = 0 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        If Shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        
        With DrawF
            .X = OX
            .Y = OY
            .Width = ow
            .Height = oh
            .Shape = Shape
            .CrashIndex = 0
        End With
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 255, 0, 0)
            GdipDrawRectangle GG, Pen, OX, OY, ow - 1, oh - 1
        End If
        
        If Debug_pos Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
        
        If ChoosePosition Then
            Call JudgeChoosePosition(OX, OY, ow, oh)
        End If
    End Sub
    Public Sub PaintArc(ByVal X As Single, ByVal Y As Single, ByVal W As Single, ByVal H As Single, degree As Single, Optional Start As Single = 0, Optional Color As Long, Optional size As Long = 1, Optional style As Integer = 0, Optional pos As PosAlign = posNormal, Optional animation As Integer = 0)
        'style:0=fill,1=border,2=sector

        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub

        If Scales <> 1 Then
            X = X * Scales: Y = Y * Scales
            W = W * Scales
            H = H * Scales
            size = size * Scales
        End If

        Dim OX As Single, OY As Single, ow As Single, oh As Single
        OX = X: OY = Y: ow = W: oh = H
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during)
                CopyMemory co(0), Color, 4
                al = co(3): al = al / 255
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, OX, OY, ow, oh, al, pro
                End If
                Color = argb(al * 255, co(0), co(1), co(2))
            End If
        End If
        
        GdipResetPath path
        
        If pos = 1 Then OX = Int(X - ow / 2): OY = Int(Y - oh / 2)
        If pos = 2 Then OX = OX - ow
        If pos = 3 Then OY = OY - oh
        If pos = 4 Then OX = OX + ow
        If pos = 5 Then OY = OY + oh
        
        GdipAddPathArc path, OX, OY, ow, oh, Start, degree
        If style = 2 Then GdipAddPathLine path, OX + ow / 2, OY + ow / 2, OX + ow / 2, OY + ow / 2
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
            .X = OX
            .Y = OY
            .Width = ow
            .Height = oh
            .CrashIndex = 0
            .Shape = 0
        End With
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 255, 255)
            GdipDrawRectangle GG, Pen, OX, OY, ow - 1, oh - 1
        End If
        
        If Debug_pos Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 183, 195)
            GdipDrawLine GG, Pen, X - 5, Y, X + 5, Y
            GdipDrawLine GG, Pen, X, Y - 5, X, Y + 5
        End If
        
        If ChoosePosition Then
            Call JudgeChoosePosition(OX, OY, ow, oh)
        End If
    End Sub
    Public Sub PaintPolygon2(Color As Long, style As Integer, Points(), Optional SmoothMode As Boolean = False)
    
        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub
        
        For I = 0 To UBound(Points)
            Points(I) = Points(I) * Scales
        Next

        GdipResetPath path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For I = 0 To UBound(Points) Step 2
            p(I / 2).X = Points(I): p(I / 2).Y = Points(I + 1)
        Next
        
        If SmoothMode Then
            GdipAddPathCurve path, p(0), UBound(p) + 1
        Else
            GdipAddPathPolygon path, p(0), UBound(p) + 1
        End If
        
        If style = 0 Then GdipSetSolidFillColor brush, Color
        If style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, size
        End If
        
        If style = 0 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, path
        End If
    End Sub
    Public Sub PaintPolygon(Color As Long, style As Integer, SmoothMode As Boolean, ParamArray Points())
        
        If ECore.FreezeAvailable And (Not IsSystem) Then Exit Sub
     
        For I = 0 To UBound(Points)
            Points(I) = Points(I) * Scales
        Next

        GdipResetPath path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For I = 0 To UBound(Points) Step 2
            p(I / 2).X = Points(I): p(I / 2).Y = Points(I + 1)
        Next
        
        If SmoothMode Then
            GdipAddPathCurve path, p(0), UBound(p) + 1
        Else
            GdipAddPathPolygon path, p(0), UBound(p) + 1
        End If
        
        If style = 0 Then GdipSetSolidFillColor brush, Color
        If style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, size
        End If
        
        If style = 0 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        
        If IsSystem Then Exit Sub
        
        If Debug_focus Then
            GdipSetPenWidth Pen, 1
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, path
        End If
        
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
    Public Function ShowEdit(Text As String, Shape As Integer, X As Single, Y As Single, W As Single, H As Single, TextColor As Long, Color As Long, HoverColor As Long, LineColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As MButtonState
        Dim m As Integer, R As RECT

        m = CheckMouse(X, Y, W - IIf(TextHandle = VarPtr(Text), H, 0), H)
        Paint Shape, X, Y, W, H, IIf(m <> 0 Or TextHandle = VarPtr(Text), HoverColor, Color), Radius, style:=ShapeStyle
        If m = 3 Then TextHandle = VarPtr(Text)
        
        Dim CtrlPressed As Boolean
        CtrlPressed = (GetAsyncKeyState(vk_ctrl) <> 0)
        
        If VarPtr(Text) = TextHandle And WaitChr <> "" Then
            If WaitChr = Chr(vbKeyReturn) Then
                TextHandle = 0
                m = 4
            ElseIf WaitChr = Chr(vbKeyBack) Then
                If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1)
            ElseIf Asc(WaitChr) = 22 Then '粘贴
                Text = Text & Clipboard.GetText
            ElseIf Asc(WaitChr) = 3 Then '复制
                Clipboard.Clear
                Clipboard.SetText Text
            ElseIf Asc(WaitChr) = 24 Then '剪切
                Clipboard.Clear
                Clipboard.SetText Text
                Text = ""
            ElseIf Asc(WaitChr) <= 26 And Asc(WaitChr) >= 0 Then '其他Ctrl组合键
                
            Else
                Text = Text & WaitChr
            End If
            WaitChr = ""
        End If
        
        Writes Text, X + Radius / 4, Y + H / 2 - size / 0.75 / 2 - 1, size, TextColor, W - Radius / 2 - H, size / 0.75, StringAlignmentNear, style
        
        If TextHandle = VarPtr(Text) Then
            Dim w2 As Long, pro As Long, Alpha As Single
            w2 = EF.GetWidth(GG, Text, size, StringAlignmentNear, style)
            If w2 > W - Radius / 2 - H Then
                If Len(Text) > 0 Then Text = Left(Text, Len(Text) - 1): VBA.Beep
            End If
            Paint Shape, X, Y, W, H, LineColor, size:=2, Radius:=Radius, style:=1
            pro = GetTickCount Mod 1000
            If pro <= 700 Then
                Alpha = 1 - Cubic(pro / 700, 0, 1, 1, 1)
            Else
                Alpha = Cubic((pro - 700) / 300, 0, 1, 1, 1)
            End If
            Dim co(3) As Byte, co2 As Long
            co2 = IIf(m <> 0 Or TextHandle = VarPtr(Text), HoverColor, Color)
            CopyMemory co(0), co2, 4
            If w2 = 0 Then w2 = size / 4
            If Alpha <> 0 Then Paint 0, X + w2 + Radius / 4 - size / 8, Y + H / 2 - size / 2, 3, size, argb(Int(Alpha * 255), 255 - co(2), 255 - co(1), 255 - co(0))
            
            Paint Shape, X + W - H, Y, H, H, LineColor, Radius:=Radius
            If CheckMouse2 = mMouseUp Then
                TextHandle = 0
                m = 4
            End If
            
            If Shape = 2 Then
                Writes ">", X + W - H, Y + H / 2 - size / 0.75 / 2 - 1, size, TextColor, H + 6, size / 0.75, StringAlignmentCenter, FontStyleBold
            Else
                Writes ">", X + W - H, Y + H / 2 - size / 0.75 / 2 - 2, size, TextColor, H + 3, size / 0.75, StringAlignmentCenter, FontStyleBold
            End If
        End If
        
        If Mouse.State = 2 And m = 0 Then TextHandle = 0: m = 4
        
        ShowEdit = m
        
    End Function
    Public Sub ShowLoading(X As Single, Y As Single, W As Single, H As Single, size As Long, color1 As Long, color2 As Long, color3 As Long)
        
        PaintArc X, Y, W, H, 120, (GetTickCount Mod 3000) / 3000 * 360, color3, size * 0.86, 1
        PaintArc X, Y, W, H, 120, (GetTickCount Mod 1800) / 1800 * 360 + 90, color2, size * 0.93, 1
        PaintArc X, Y, W, H, 120, (GetTickCount Mod 1200) / 1200 * 360 + 90, color1, size, 1
        'argb(255, 242, 195, 17),argb(255, 240, 127, 34),argb(255, 232, 76, 61)
    End Sub
    Public Function ShowButton(pic As String, X As Long, Y As Long, Text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As MButtonState

        Dim W As Single, H As Single, I As Integer, m As Integer
        I = Res.GetImage(pic)
        W = Res.ImgSize(I, imgGetWidth): H = Res.ImgSize(I, imgGetHeight)
        
        m = CheckMouse(X, Y, W, Int(H / 2))
        DrawImage I, X, Y, CY:=IIf(m <> 0, Int(H / 2), 0), ch:=Int(H / 2), Alpha:=1
        
        Writes Text, X, Y + H / 2 / 2 - EF.EmHeight / 18 * size / 2 + 1, size, Color, W, size / 0.75, StringAlignmentCenter, style
        
        ShowButton = m
    End Function
    Public Function ShowSimpleButton(pic As String, X As Long, Y As Long, Text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As MButtonState

        Dim W As Single, H As Single, I As Integer, m As Integer
        I = Res.GetImage(pic)
        W = Res.ImgSize(I, imgGetWidth): H = Res.ImgSize(I, imgGetHeight)
        
        m = CheckMouse(X, Y, W, H)
        DrawImage I, X, Y, Alpha:=IIf(m, 1, 0.8)
        
        Writes Text, X, Y + H / 2 - EF.EmHeight / 18 * size / 2 + 1, size, Color, W, size / 0.75, StringAlignmentCenter, style
    
        ShowSimpleButton = m
    End Function
    Public Function ShowColorButton(Shape As Integer, X As Single, Y As Single, W As Single, H As Single, Text As String, TextColor As Long, Color As Long, HoverColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As MButtonState
        Dim m As Integer

        m = CheckMouse(X, Y, W, H)
        Paint Shape, X, Y, W, H, IIf(m, HoverColor, Color), Radius, style:=ShapeStyle
        
        Writes Text, X, Y + H / 2 - EF.EmHeight / 18 * size / 2 + 1, size, TextColor, W, size / 0.75, StringAlignmentCenter, style
    
        ShowColorButton = m
    End Function
    Public Function ShowCheckBox(value, pic As String, X As Long, Y As Long, Text As String, Color As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As MButtonState

        Dim W As Single, H As Single, I As Integer, m As Integer
        Dim X2 As Long, Y2 As Long
        X2 = X: Y2 = Y
        
        I = Res.GetImage(pic)
        W = Res.ImgSize(I, imgGetWidth): H = Res.ImgSize(I, imgGetHeight)
        
        m = CheckMouse(X, Y, W, Int(H / 2))
        DrawImage I, X, Y, CY:=IIf(value = True, Int(H / 2), 0), ch:=Int(H / 2), Alpha:=1
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes Text, X2 + W + EF.EmHeight / 18 * size / 2 + 1, Y2 + H / 2 / 2 - EF.EmHeight / 18 * size / 2 + 1, size, Color, GW, size / 0.75, StringAlignmentNear, style
        
        ShowCheckBox = m
    End Function
    Public Function ShowColorCheckBox(value, X As Single, Y As Single, W As Single, H As Single, Text As String, DefaultColor As Long, ChooseColor As Long, ForeColor As Long, Optional size As Long = 14, Optional style As FontStyle = FontStyleRegular) As MButtonState

        m = CheckMouse(X, Y, W, H)
        
        Paint 1, X, Y, H, H, IIf(value = True, ChooseColor, DefaultColor)
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes Text, X + H + EF.EmHeight / 18 * size / 2 + 1, Y + H / 2 - EF.EmHeight / 18 * size / 2 + 1, size, ForeColor, GW, size / 0.75, StringAlignmentNear, style
        
        ShowColorCheckBox = m
    End Function
    Public Function ShowSliderBar(value, pic As String, X As Long, Y As Long) As MButtonState

        Dim W As Single, H As Single, I As Integer, m As Integer
        Dim X2 As Long, Y2 As Long
        
        X2 = X: Y2 = Y
        
        I = Res.GetImage(pic)
        W = Res.ImgSize(I, imgGetWidth): H = Res.ImgSize(I, imgGetHeight)
        
        m = CheckMouse(X, Y, W, Int(H / 2))
        DrawImage I, X, Y, CY:=0, ch:=Int(H / 2), Alpha:=1
        DrawImage I, X, Y, CY:=Int(H / 2), cw:=value * W, ch:=Int(H / 2), Alpha:=1
        
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.X - X) / W
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowSliderBar = m
    End Function
    Public Function ShowHScrollBar(value, X As Single, Y As Single, W As Single, H As Single, LineColor As Long, BallColor As Long, Optional size As Long = 2) As MButtonState
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
    Public Function ShowVScrollBar(value, X As Single, Y As Single, W As Single, H As Single, LineColor As Long, BallColor As Long, Optional size As Long = 2) As MButtonState
        Dim m As Integer

        m = CheckMouse(X, Y, W, H)
        PaintLine X + W / 2 - size / 2, Y, X + W / 2 - size / 2, Y + H, LineColor, size
        Paint 1, X, Y + value * (H - W), W, W, BallColor
        
        If m > 0 And Mouse.button > 0 Then
            value = (Mouse.Y - Y) / (H - W)
            If value < 0 Then value = 0
            If value > 1 Then value = 1
        End If
        
        ShowVScrollBar = m
    End Function
'========================================================


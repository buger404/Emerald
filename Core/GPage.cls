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

Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As clsid) As Long
Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
Private Type GDA
    x As Long
    y As Long
    w As Long
    h As Long
End Type
Private Type AniTask
    Start As Long
    during As Long
    custom As Boolean
    func As String
    profunc As Integer
    delay As Long
    mark As Boolean
End Type
Private Const GTick As Long = 1000 / 30
Dim Anis() As AniTask
Dim Mems() As GMem, DrawBox As Object, brush As Long, Path As Long, Pen As Long
Dim ScrollMode As Boolean, ScrollX As Long, ScrollY As Long, ScrollBX As Long, ScrollBY As Long, CRgn As Long
Public CDC As Long, GG As Long, OutOfScroll As Boolean
'========================================================
'   Init
    Public Sub Create(dBox As Object)
        ReDim Mems(0): ReDim Anis(0)
        
        CDC = CreateCDC(GW, GH)
        NewGM CDC, 0, ""
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
        For i = 1 To UBound(Mems)
            If Not Mems(i).copyed Then
                Select Case Mems(i).kind
                    Case 0 'DC
                        DeleteObject Mems(i).Hwnd
                        GdipDisposeImage Mems(i).ImgHwnd
                    Case 1 'Graphics
                        GdipDeleteGraphics Mems(i).Hwnd
                End Select
                If Mems(i).GIF.Count > 0 Then
                    For s = 0 To Mems(i).GIF.Count
                        DeleteObject Mems(i).GIF.frames(s)
                    Next
                End If
            End If
        Next
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
    Private Sub NewGM(mem As Long, kind As Integer, name As String, Optional w, Optional h, Optional folder As String = "")
        ReDim Preserve Mems(UBound(Mems) + 1)
        With Mems(UBound(Mems))
            .Hwnd = mem
            .kind = kind
            .name = name
            .folder = folder
            If Not IsMissing(w) Then .w = w
            If Not IsMissing(h) Then .h = h
        End With
    End Sub
    Public Property Get SrcClass() As Object
        Set SrcClass = DrawBox
    End Property
    Public Property Get ImgCount() As Long
        ImgCount = UBound(Mems)
    End Property
    Public Property Get ImgSize(n, sizeIndex As imgIndex)
        Dim index As Integer
        
        If TypeName(n) = "String" Then
            index = GetImage(n)
        Else
            index = n
        End If
    
        Select Case sizeIndex
            Case 0
                ImgSize = Mems(index).w
            Case 1
                ImgSize = Mems(index).h
            Case 2
                ImgSize = Mems(index).GIF.Count
        End Select
    End Property
'========================================================
'   Images
    Public Sub NewImages(Path As String, Optional arg1, Optional arg2)
        Dim f As String, c As Long
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        
        Dim index As Integer, Tree As AssetsTree
        Tree.Path = Path: ReDim Tree.Files(0)
        index = FindAssetsTree(Path)
        
        f = Dir(Path)
        Do While f <> ""
            c = c + 1
            ReDim Preserve Tree.Files(UBound(Tree.Files) + 1)
            f = Dir()
        Loop
        If index = 0 Then
            PreLoadCount = PreLoadCount + c
        Else
            Tree = GetAssetsTree(Path)
            ReLoadCount = ReLoadCount + UBound(Tree.Files)
            For i = 1 To UBound(Tree.Files)
                ReDim Preserve Mems(UBound(Mems) + 1)
                Mems(UBound(Mems)) = Tree.Files(i)
                Mems(UBound(Mems)).copyed = True
            Next
            Exit Sub
        End If
        
        f = Dir(Path)
        i = 1
        Do While f <> ""
            If IsMissing(arg1) Then
                newImage Path & f
            ElseIf IsMissing(arg2) Then
                newImage Path & f, arg1
            Else
                newImage Path & f, arg1, arg2
            End If
            Tree.Files(i) = Mems(UBound(Mems))
            i = i + 1
            f = Dir()
            DoEvents
            LoadedCount = LoadedCount + 1
        Loop
        
        If index = 0 Then AddAssetsTree Tree
    End Sub
    Public Sub newImage(Path As String, Optional arg1, Optional arg2)
        'when arg1 is missing ; i.w = i.w
        'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
        'other ; i.w = arg1
        Dim si As Integer, i As Long, w As Long, h As Long, DC As Long, g As Long
        If IsMissing(arg1) Then si = 0
        If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
        If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
        
        GdipCreateBitmapFromFile StrPtr(Path), i
        GdipGetImageWidth i, w: GdipGetImageHeight i, h
        
        If si = 1 Then w = Round(w * arg1): h = Round(h * arg1)
        If si = 2 Then w = arg1: h = arg2
        
        Dim gC As clsid, gL As Long, index As Long, temp() As String
        CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
        GdipImageGetFrameCount i, gC, gL
        If gL > 0 Then  '动态GIF支持
            temp = Split(Path, "\")
            NewGM 0, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
            With Mems(UBound(Mems))
                .GIF.Count = gL
                ReDim .GIF.frames(gL)
                For index = 0 To gL
                    GdipImageSelectActiveFrame i, gC, index
                    DC = CreateCDC(w, h)
                    GdipCreateFromHDC DC, g
                    GdipDrawImageRect g, i, 0, 0, w, h
                    GdipDeleteGraphics g
                    .GIF.frames(index) = DC
                Next
            End With
            
            Mems(UBound(Mems)).ImgHwnd = i
            Exit Sub
        End If
        
        DC = CreateCDC(w, h)
        GdipCreateFromHDC DC, g
        GdipDrawImageRect g, i, 0, 0, w, h
        
        GdipDeleteGraphics g
        
        temp = Split(Path, "\")
        NewGM DC, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
        Mems(UBound(Mems)).ImgHwnd = i
    End Sub
    Public Sub DrawImageEx(n, x As Long, y As Long, Optional w, Optional h, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
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
            index = GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(w) Then GdipGetImageWidth Mems(index).ImgHwnd, w
        If IsMissing(h) Then GdipGetImageWidth Mems(index).ImgHwnd, h
        
        If Pos = 1 Then x = Int(x - w / 2): y = Int(y - h / 2)
        If Pos = 2 Then x = x - w
        If Pos = 3 Then y = y - h
        If Pos = 4 Then x = x + w
        If Pos = 5 Then y = y + h
        
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long
            GdipIsVisibleRect GG, x - 1, y - 1, w, h, Ret
            If Not Ret Then OutOfScroll = True: Exit Sub
        End If
        
        GdipDrawImageRect GG, Mems(index).ImgHwnd, x, y, w, h
        
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
    Public Sub DrawImage(n, x As Long, y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional Alpha, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
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
            index = GetImage(n)
        Else
            index = n
        End If
        
        If IsMissing(cx) Then cx = 0
        If IsMissing(cy) Then cy = 0
        If IsMissing(cw) Then cw = Mems(index).w - cx
        If IsMissing(ch) Then ch = Mems(index).h - cy
        
        If Pos = 1 Then x = Int(x - cw / 2): y = Int(y - ch / 2)
        If Pos = 2 Then x = x - cw
        If Pos = 3 Then y = y - ch
        If Pos = 4 Then x = x + cw
        If Pos = 5 Then y = y + ch
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, ((GetTickCount - Anis(animation).Start - Anis(animation).delay) / Anis(animation).during))
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, cw, ch, Alpha, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, cw, ch, Alpha, pro
                End If
                
                GoSub setAlpha
            End If
        End If
        
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim Ret As Long, Crect As RECT
            With Crect
                .Left = x: .top = y: .Right = x + cw: .Bottom = y + ch
            End With
            If RectVisible(CDC, Crect) <> 1 Then OutOfScroll = True: Exit Sub
        End If
        
        Dim srcDC As Long
        If Mems(index).GIF.Count = 0 Then
            srcDC = Mems(index).Hwnd
        Else
            If GetTickCount - Mems(index).GIF.time >= GTick Then Mems(index).GIF.tick = Mems(index).GIF.tick + 1
            If Mems(index).GIF.tick > Mems(index).GIF.Count Then Mems(index).GIF.tick = 0
            srcDC = Mems(index).GIF.frames(Mems(index).GIF.tick)
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
    Private Function GetImage(ByVal name As String) As Integer
        For i = 1 To UBound(Mems)
            If Mems(i).kind = 0 Then
                If Mems(i).name = name Then GetImage = i: Exit For
            End If
        Next
    End Function
'========================================================
'   Writer
    Public Sub Writes(ByVal text As String, x As Long, y As Long, Optional Size As Long = 14, Optional Color As Long, Optional w As Long = 0, Optional h As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional Style As FontStyle = FontStyleRegular, Optional animation As Integer = 0)
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
        
        EF.Writes text, x, y, GG, Color, Size, w, h, align, Style
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
    Public Sub Paint(ByVal shape As Integer, x As Long, y As Long, w As Long, h As Long, Optional Color As Long, Optional Radius As Long, Optional Size As Long = 1, Optional Style As Integer = 0, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
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
        
        GdipResetPath Path
        
        If Pos = 1 Then x = Int(x - w / 2): y = Int(y - h / 2)
        If Pos = 2 Then x = x - w
        If Pos = 3 Then y = y - h
        If Pos = 4 Then x = x + w
        If Pos = 5 Then y = y + h
        
ReShape:
        If shape = 0 Then GdipAddPathRectangle Path, x, y, w - 1, h - 1
        If shape = 1 Then GdipAddPathEllipse Path, x, y, w - 1, h - 1
        If shape = 2 Then
        
            If Radius = 0 Then
                shape = 0: GoTo ReShape
            End If
            
            If Radius > w Then Radius = w
            If Radius > h Then Radius = h
            
            GdipAddPathArc Path, x, y, Radius, Radius, 180, 90
            GdipAddPathArc Path, x + w - Radius, y, Radius, Radius, 270, 90
            GdipAddPathArc Path, x + w - Radius, y + h - Radius, Radius, Radius, 0, 90
            GdipAddPathArc Path, x, y + h - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure Path
            
        End If
        
        If Color <> 0 Then
            If Style = 0 Then GdipSetSolidFillColor brush, Color
            If Style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, Size
            End If
        End If
        
        If shape = 0 Then GdipSetSmoothingMode GG, SmoothingModeDefault
        If Style = 0 Then GdipFillPath GG, brush, Path
        If Style = 1 Then GdipDrawPath GG, Pen, Path
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
    Public Sub PaintArc(x As Long, y As Long, w As Long, h As Long, Degree As Long, Optional Start As Long = 0, Optional Color As Long, Optional Size As Long = 1, Optional Style As Integer = 0, Optional Pos As PosAlign = posNormal, Optional animation As Integer = 0)
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
        
        GdipResetPath Path
        
        If Pos = 1 Then x = Int(x - w / 2): y = Int(y - h / 2)
        If Pos = 2 Then x = x - w
        If Pos = 3 Then y = y - h
        If Pos = 4 Then x = x + w
        If Pos = 5 Then y = y + h
        
        GdipAddPathArc Path, x, y, w, h, Start, Degree
        If Style = 2 Then GdipAddPathLine Path, x + w / 2, y + w / 2, x + w / 2, y + w / 2
        If Style = 0 Then GdipClosePathFigure Path
        
        If Color <> 0 Then
            If Style = 0 Or Style = 2 Then GdipSetSolidFillColor brush, Color
            If Style = 1 Then
                GdipSetPenColor Pen, Color
                GdipSetPenWidth Pen, Size
            End If
        End If
        
        If Style = 0 Or Style = 2 Then GdipFillPath GG, brush, Path
        If Style = 1 Then GdipDrawPath GG, Pen, Path
        
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
    Public Sub PaintPolygon(Color As Long, Style As Integer, ParamArray Points())

        GdipResetPath Path
        Dim p() As POINTF
        ReDim p((UBound(Points) - 1) / 2)
        For i = 0 To UBound(Points) Step 2
            p(i / 2).x = Points(i): p(i / 2).y = Points(i + 1)
        Next
        
        GdipAddPathPolygon Path, p(0), UBound(p) + 1
        
        If Style = 0 Then GdipSetSolidFillColor brush, Color
        If Style = 1 Then
            GdipSetPenColor Pen, Color
            GdipSetPenWidth Pen, Size
        End If
        
        If Style = 0 Then GdipFillPath GG, brush, Path
        If Style = 1 Then GdipDrawPath GG, Pen, Path
        
        If Debug_focus Then
            GdipSetPenColor Pen, argb(255, 0, 176, 240)
            GdipDrawPath GG, Pen, Path
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
    Public Function ShowButton(pic As String, x As Long, y As Long, text As String, Color As Long, Optional Size As Long = 14, Optional Style As FontStyle = FontStyleRegular) As Integer
        Dim w As Long, h As Long, i As Integer, m As Integer
        i = GetImage(pic)
        w = Mems(i).w: h = Mems(i).h
        
        m = CheckMouse(x, y, w, Int(h / 2))
        DrawImage i, x, y, cy:=IIf(m <> 0, Int(h / 2), 0), ch:=Int(h / 2), Alpha:=1
        
        Writes text, x, y + h / 2 / 2 - Size / 0.75 / 2, Size, Color, w, Size / 0.75, StringAlignmentCenter, Style
        
        ShowButton = m
    End Function
    Public Function ShowSimpleButton(pic As String, x As Long, y As Long, text As String, Color As Long, Optional Size As Long = 14, Optional Style As FontStyle = FontStyleRegular) As Integer
        Dim w As Long, h As Long, i As Integer, m As Integer
        i = GetImage(pic)
        w = Mems(i).w: h = Mems(i).h
        
        m = CheckMouse(x, y, w, h)
        DrawImage i, x, y, Alpha:=IIf(m, 1, 0.8)
        
        Writes text, x, y + h / 2 - Size / 0.75 / 2, Size, Color, w, Size / 0.75, StringAlignmentCenter, Style
    
        ShowSimpleButton = m
    End Function
    Public Function ShowColorButton(shape As Integer, x As Long, y As Long, w As Long, h As Long, text As String, TextColor As Long, Color As Long, HoverColor As Long, Optional Radius As Long = 0, Optional ShapeStyle As Integer = 0, Optional Size As Long = 14, Optional Style As FontStyle = FontStyleRegular) As Integer
        Dim m As Integer
        
        m = CheckMouse(x, y, w, h)
        Paint shape, x, y, w, h, IIf(m, HoverColor, Color), Radius, Style:=ShapeStyle
        
        Writes text, x, y + h / 2 - Size / 0.75 / 2, Size, TextColor, w, Size / 0.75, StringAlignmentCenter, Style
    
        ShowColorButton = m
    End Function
    Public Function ShowCheckBox(value, pic As String, x As Long, y As Long, text As String, Color As Long, Optional Size As Long = 14, Optional Style As FontStyle = FontStyleRegular) As Integer
        Dim w As Long, h As Long, i As Integer, m As Integer
        i = GetImage(pic)
        w = Mems(i).w: h = Mems(i).h
        
        m = CheckMouse(x, y, w, Int(h / 2))
        DrawImage i, x, y, cy:=IIf(value = True, Int(h / 2), 0), ch:=Int(h / 2), Alpha:=1
        If m = 3 Then value = IIf(value = True, False, True)
        
        Writes text, x + w + Size / 0.75 / 2, y + h / 2 / 2 - Size / 0.75 / 2, Size, Color, GW, Size / 0.75, StringAlignmentNear, Style
        
        ShowCheckBox = m
    End Function
'========================================================

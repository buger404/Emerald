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
    Count As Long
End Type
Private Type GMem
    GIF As GGIF
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
Private Type AniTask
    start As Long
    during As Long
    custom As Boolean
    func As String
    profunc As Integer
    delay As Long
    mark As Boolean
End Type
Private Const GTick As Long = 1000 / 30
Dim Anis() As AniTask
Dim Mems() As GMem, DrawBox As Object, brush As Long, path As Long, Pen As Long
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
            If Mems(i).GIF.Count > 0 Then
                For s = 0 To Mems(i).GIF.Count
                    DeleteObject Mems(i).GIF.frames(s)
                Next
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
                If (GetTickCount - Anis(i).start - Anis(i).delay) > Anis(i).during Then
                     CallByName DrawBox, "AnimationDone", VbMethod, i
                    Anis(i).mark = True
                End If
            End If
        Next
        
        CallByName DrawBox, "Update", VbMethod
    End Sub
    Public Sub Display(dc As Long)
        BitBlt dc, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
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
    Public Property Get SrcClass() As Object
        Set SrcClass = DrawBox
    End Property
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
        Dim si As Integer, i As Long, w As Long, h As Long, dc As Long, g As Long
        If IsMissing(arg1) Then si = 0
        If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
        If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
        
        GdipCreateBitmapFromFile StrPtr(path), i
        GdipGetImageWidth i, w: GdipGetImageHeight i, h
        
        If si = 1 Then w = Round(w * arg1): h = Round(h * arg1)
        If si = 2 Then w = arg1: h = arg2
        
        Dim gC As clsid, gL As Long, index As Long, temp() As String
        CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
        GdipImageGetFrameCount i, gC, gL
        If gL > 0 Then  '¶¯Ì¬GIFÖ§³Ö
            temp = Split(path, "\")
            NewGM 0, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
            With Mems(UBound(Mems))
                .GIF.Count = gL
                ReDim .GIF.frames(gL)
                For index = 0 To gL
                    GdipImageSelectActiveFrame i, gC, index
                    dc = CreateCDC(w, h)
                    GdipCreateFromHDC dc, g
                    GdipDrawImageRect g, i, 0, 0, w, h
                    GdipDeleteGraphics g
                    .GIF.frames(index) = dc
                Next
            End With
            
            GdipDisposeImage i
            Exit Sub
        End If
        
        dc = CreateCDC(w, h)
        GdipCreateFromHDC dc, g
        GdipDrawImageRect g, i, 0, 0, w, h
        
        GdipDisposeImage i: GdipDeleteGraphics g
        
        temp = Split(path, "\")
        NewGM dc, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
    End Sub
    Public Sub DrawImage(n, x As Long, y As Long, Optional cx, Optional cy, Optional cw, Optional ch, Optional alpha, Optional Pos As Integer = 0, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        
        If animation <> 0 Then
            If (GetTickCount - Anis(animation).start) > Anis(animation).during Then animation = 0
        End If
        
        If Not IsMissing(alpha) Then
            If animation = 0 Then
setAlpha:
                If alpha < 0 Then alpha = 0
                If alpha > 1 Then alpha = 1
                With b
                    .AlphaFormat = &H1
                    .BlendFlags = &H0
                    .BlendOp = 0
                    .SourceConstantAlpha = Int(alpha * 255)
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
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, ((GetTickCount - Anis(animation).start - Anis(animation).delay) / Anis(animation).during))
                
                If Anis(animation).custom Then
                    CallByName DrawBox, Anis(animation).func, VbMethod, x, y, cw, ch, alpha, pro
                Else
                    CallByName EAni, Anis(animation).func, VbMethod, x, y, cw, ch, alpha, pro
                End If
                
                GoSub setAlpha
            End If
        End If
        
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim ret As Long, Crect As RECT
            With Crect
                .Left = x: .top = y: .Right = x + cw: .Bottom = y + ch
            End With
            If RectVisible(CDC, Crect) <> 1 Then OutOfScroll = True: Exit Sub
        End If
        
        Dim srcDC As Long
        If Mems(index).GIF.Count = 0 Then
            srcDC = Mems(index).hwnd
        Else
            If GetTickCount - Mems(index).GIF.time >= GTick Then Mems(index).GIF.tick = Mems(index).GIF.tick + 1
            If Mems(index).GIF.tick > Mems(index).GIF.Count Then Mems(index).GIF.tick = 0
            srcDC = Mems(index).GIF.frames(Mems(index).GIF.tick)
        End If
        
        If IsMissing(alpha) Then
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
    Public Sub Writes(ByVal text As String, x As Long, y As Long, Optional size As Long = 14, Optional Color As Long, Optional w As Long = 0, Optional h As Long = 0, Optional align As StringAlignment = StringAlignmentNear, Optional style As FontStyle = FontStyleRegular, Optional animation As Integer = 0)
        'If OutOfScroll Then Exit Sub
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim ret As Long
            GdipIsVisibleRect GG, x, y, w, h, ret
            If ret = 0 Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).start - Anis(animation).delay) / Anis(animation).during)
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
    End Sub
'========================================================
'   Shape
    Public Sub Paint(ByVal shape As Integer, x As Long, y As Long, w As Long, h As Long, Optional Color As Long, Optional Radius As Long, Optional style As Integer = 0, Optional animation As Integer = 0)
        'shape:0=rect,1=ellipse,2=rectr
        'style:0=fill,1=border
        'If OutOfScroll Then Exit Sub
        If ScrollMode Then
            x = x + ScrollX + ScrollBX: y = y + ScrollY + ScrollBY
            Dim ret As Long
            GdipIsVisibleRect GG, x - 1, y - 1, w, h, ret
            If Not ret Then OutOfScroll = True: Exit Sub
        End If
        
        If animation <> 0 Then
            If Not Anis(animation).mark Then
                Dim pro As Single, co(3) As Byte, al As Single
                pro = CallByName(EAni, "GetProgress_" & Anis(animation).profunc, VbMethod, (GetTickCount - Anis(animation).start - Anis(animation).delay) / Anis(animation).during)
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
        If shape = 0 Then GdipAddPathRectangle path, x - 1, y - 1, w, h
        If shape = 1 Then GdipAddPathEllipse path, x, y, w - 1, h - 1
        
        If Color <> 0 Then
            If style = 0 Then GdipSetSolidFillColor brush, Color
            If style = 1 Then GdipSetPenColor Pen, Color
        End If
        
        If style = 0 Then GdipFillPath GG, brush, path
        If style = 1 Then GdipDrawPath GG, Pen, path
        
        With DrawF
            .Left = x
            .top = y
            .Right = w
            .Bottom = h
        End With
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
        Anis(id).start = GetTickCount
        Anis(id).delay = delay
        Anis(id).mark = False
    End Sub
    Public Function AnimationDone(id As Integer) As Boolean
        AnimationDone = Anis(id).mark
    End Function
'========================================================

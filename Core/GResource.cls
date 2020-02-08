VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As clsid) As Long
Public HotLoad As Boolean
Dim AssetsFolder() As String, FolderParam1() As Variant, FolderParam2() As Variant
Public SearchIndex As Long
Public NoLoadCrash As Boolean, NoLoadDirection As Boolean
Dim Mems() As GMem
Dim Effect As Long, effectw As Long, effecth As Long, effectimg As Long, effectindex As Integer
Dim DisposeMark As Boolean
Private Sub Class_Initialize()
    ReDim Mems(0)
    ReDim AssetsFolder(0): ReDim FolderParam1(0): ReDim FolderParam2(0)
    SearchIndex = 1
    NoLoadDirection = True
End Sub
Public Sub ClipCircle(n)
    Dim index As Integer
    
    If TypeName(n) = "String" Then
        index = GetImage(n)
    Else
        index = n
    End If

    Dim path As Long, bmpGraph As Long, BMP As Long, BMP2 As Long, tempGraph As Long
    Dim w As Long, h As Long
    
    GdipGetImageWidth Mems(index).ImgHwnd, w
    GdipGetImageHeight Mems(index).ImgHwnd, h

    GdipCreateBitmapFromScan0 w, h, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP2
    GdipCreateBitmapFromScan0 w * 2, h * 2, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP

    GdipGetImageGraphicsContext BMP2, tempGraph
    GdipGetImageGraphicsContext BMP, bmpGraph
    GdipSetSmoothingMode tempGraph, SmoothingModeAntiAlias
    GdipSetSmoothingMode bmpGraph, SmoothingModeAntiAlias
    
    PoolCreatePath FillModeWinding, path
    GdipAddPathEllipseI path, 1, 1, w * 2 - 2, h * 2 - 2
    GdipSetClipPath bmpGraph, path, CombineModeReplace
    
    GdipDrawImageRectI bmpGraph, Mems(index).ImgHwnd, 0, 0, w * 2, h * 2
    GdipResetClip bmpGraph
    
    GdipSetInterpolationMode tempGraph, InterpolationModeHighQualityBilinear
    
    GdipDrawImageRectI tempGraph, BMP, 0, 0, w - 1, h - 1

    GdipDeletePath path
    PoolDeleteGraphics bmpGraph
    PoolDisposeImage BMP
    PoolDeleteGraphics tempGraph
    PoolDisposeImage Mems(index).ImgHwnd
    
    Mems(index).ImgHwnd = BMP2
    
    Call ReDraw(index)
End Sub
Private Sub ApplyEffect(n, Kind As GdipEffectType)
    If XPMode Then Exit Sub

    Dim index As Integer
    
    If TypeName(n) = "String" Then
        index = GetImage(n)
    Else
        index = n
    End If
    
    effectimg = Mems(index).ImgHwnd
    
    PoolCreateEffect2 Kind, Effect
    
    effectindex = index
End Sub
Private Sub ApplyEffectStep2()
    If XPMode Then Exit Sub

    GdipGetImageWidth effectimg, effectw: GdipGetImageHeight effectimg, effecth
    GdipBitmapApplyEffect effectimg, Effect, NewRectL(0, 0, effectw, effecth), 0, 0, 0
    
    PoolDeleteEffect Effect '垃圾处理
    
    Call ReDraw(effectindex)
End Sub
Public Sub ApplyTintEffect(n, amount As Long, hue As Long)
    If XPMode Then Exit Sub

    Dim e As TintParams
    e.amount = amount
    e.hue = hue
    
    ApplyEffect n, Tint
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplySharpenEffect(n, amount As Single, Radius As Single)
    If XPMode Then Exit Sub

    Dim e As SharpenParams
    e.amount = amount
    e.Radius = Radius

    ApplyEffect n, Sharpen
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBlurEffect(n, Radius As Long, edge As Integer)
    If XPMode Then Exit Sub

    Dim e As BlurParams
    e.Radius = Radius
    e.expandEdge = edge

    ApplyEffect n, Blur
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBrightnessContrastEffect(n, brightnessLevel As Long, contrastLevel As Long)
    If XPMode Then Exit Sub

    Dim e As BrightnessContrastParams
    e.brightnessLevel = brightnessLevel
    e.contrastLevel = contrastLevel

    ApplyEffect n, BrightnessContrast
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorBalanceEffect(n, cyanRed As Long, magentaGreen As Long, yellowBlue As Long)
    If XPMode Then Exit Sub

    Dim e As ColorBalanceParams
    e.cyanRed = cyanRed
    e.magentaGreen = magentaGreen
    e.yellowBlue = YellowGreen

    ApplyEffect n, ColorBalance
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorCurveEffect(n, adjustment As CurveAdjustments, adjustValue As Long, channel As CurveChannel)
    If XPMode Then Exit Sub

    Dim e As ColorCurveParams
    e.adjustment = adjustment
    e.adjustValue = adjustValue
    e.channel = channel

    ApplyEffect n, ColorCurve
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorMatrixEffect(n, m() As Single)
    If XPMode Then Exit Sub

    Dim e As ColorMatrix
    For I = 0 To 4
        For S = 0 To 4
            e.m(I, S) = m(I, S)
        Next
    Next
    ApplyEffect n, ColorMatrix
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyHueSaturationLightnessEffect(n, hueLevel As Long, lightnessLevel As Long, saturationLevel As Long)
    If XPMode Then Exit Sub

    Dim e As HueSaturationLightnessParams
    e.hueLevel = hueLevel
    e.lightnessLevel = lightnessLevel
    e.saturationLevel = saturationLevel

    ApplyEffect n, HueSaturationLightness
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyLevelsEffect(n, highlight As Long, midtone As Long, Shadow As Long)
    If XPMode Then Exit Sub

    Dim e As LevelsParams
    e.highlight = highlight
    e.midtone = midtone
    e.Shadow = Shadow

    ApplyEffect n, levels
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Private Sub ReDraw(I As Integer)
    If Mems(I).GIF.Count > 0 Then Exit Sub
    
    Dim g As Long
    PoolCreateFromHdc Mems(I).Hwnd, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipGraphicsClear g, 0
    GdipDrawImageRect g, Mems(I).ImgHwnd, 0, 0, Mems(I).w, Mems(I).h
    PoolDeleteGraphics g
    
    For S = 0 To 3
        PoolCreateFromHdc Mems(I).Imgs(S), g
        GdipSetSmoothingMode g, SmoothingModeAntiAlias
        GdipGraphicsClear g, 0
        GdipDrawImageRect g, Mems(I).ImgHwnd, IIf(S = 1 Or S = 3, Mems(I).w, 0), IIf(S = 2 Or S = 3, Mems(I).h, 0), Mems(I).w * IIf(S = 1 Or S = 3, -1, 1), Mems(I).h * IIf(S = 2 Or S = 3, -1, 1)
        PoolDeleteGraphics g
    Next
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim iW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(Mems(I).w + 1, Mems(I).h + 1)
    
    DC = Mems(I).Imgs(0)
    GdipGetImageWidth Mems(I).ImgHwnd, iW: GdipGetImageHeight Mems(I).ImgHwnd, iH
    For X = 0 To Mems(I).w - 1
        For y = 0 To Mems(I).h - 1
            '取得像素点颜色
            GdipBitmapGetPixel Mems(I).ImgHwnd, X / (Mems(I).w - 1) * iW, y / (Mems(I).h - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4
            
            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(X, y) = True
        Next
        DoEvents
    Next
    
    Mems(I).CrashIndex = UBound(ColorLists)
End Sub
Public Sub Add(mem As Long, Kind As Integer, Name As String, Optional w, Optional h, Optional Folder As String = "")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Hwnd = mem
        .Kind = Kind
        .Name = Name
        .Folder = Folder
        If Not IsMissing(w) Then .w = w
        If Not IsMissing(h) Then .h = h
    End With
End Sub
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
Public Sub Dispose()
    If DisposeMark Then Exit Sub
    Exit Sub
    For I = 1 To UBound(Mems)
        If Not Mems(I).copyed Then
            Select Case Mems(I).Kind
                Case 0 'DC
                    If Mems(I).Hwnd <> 0 Then DeleteObject Mems(I).Hwnd
                    If Mems(I).ImgHwnd <> 0 Then PoolDisposeImage Mems(I).ImgHwnd
                    For S = 0 To 3
                        If Mems(I).Imgs(S) <> 0 Then DeleteObject Mems(I).Imgs(S)
                    Next
                Case 1 'Graphics
                    PoolDeleteGraphics Mems(I).Hwnd
            End Select
            If Mems(I).GIF.Count > 0 Then
                For S = 0 To Mems(I).GIF.Count
                    DeleteObject Mems(I).GIF.frames(S)
                Next
            End If
        End If
    Next
    
    DisposeMark = True
End Sub
Public Property Get ResourceGIFHandle(index As Integer, frame As Integer) As Long
    ResourceGIFHandle = Mems(index).GIF.frames(frame)
End Property
Public Property Get ResourceGIFCurrentHandle(index As Integer) As Long
    ResourceGIFCurrentHandle = Mems(index).GIF.frames(Mems(index).GIF.tick)
End Property
Public Property Get ResourceGIFTick(index As Integer) As Long
    ResourceGIFTick = Mems(index).GIF.tick
End Property
Public Sub GIFTick(index As Integer)
    Mems(index).GIF.tick = Mems(index).GIF.tick + 1
    If Mems(index).GIF.tick > Mems(index).GIF.Count Then Mems(index).GIF.tick = 0
End Sub
Public Property Get ResourceGIFCount(index As Integer) As Long
    ResourceGIFCount = Mems(index).GIF.Count
End Property
Public Property Get ResourceGIFTime(index As Integer) As Long
    ResourceGIFTime = Mems(index).GIF.Time
End Property
Public Property Get ResourceName(index As Integer) As Long
    ResourceName = Mems(index).Name
End Property
Public Property Get ResourceHandle(index As Integer) As Long
    ResourceHandle = Mems(index).Hwnd
End Property
Public Property Get ResourceCrashIndex(index As Integer) As Long
    ResourceCrashIndex = Mems(index).CrashIndex
End Property
Public Property Get ResourceHandle2(index As Integer, Direction As ImgDirection) As Long
    ResourceHandle2 = Mems(index).Imgs(Direction)
End Property
Public Property Get ResourceImageHandle(index As Integer) As Long
    ResourceImageHandle = Mems(index).ImgHwnd
End Property
Public Function GetImage(ByVal Name As String) As Integer
    For I = SearchIndex To UBound(Mems)
        If Mems(I).Kind = 0 Then
            If Mems(I).Name = Name Then GetImage = I: Exit For
        End If
    Next
    
    If HotLoad Then
        If GetImage = 0 Then
            For I = 1 To UBound(AssetsFolder)
                If IsExitAFile(AssetsFolder(I) & Name) Then
                    newImage AssetsFolder(I) & Name, FolderParam1(I), FolderParam2(I)
                    GetImage = UBound(Mems)
                    Exit For
                End If
            Next
        End If
    End If
End Function
Public Sub newImage(path As String, Optional arg1, Optional arg2, Optional Name As String = "")
    'when arg1 is missing ; i.w = i.w
    'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
    'other ; i.w = arg1
    Dim si As Integer, I As Long, w As Long, h As Long, DC As Long, g As Long
    If IsMissing(arg1) Then si = 0
    If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
    If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
    
    PoolCreateBitmapFromFile StrPtr(path), I
    GdipGetImageWidth I, w: GdipGetImageHeight I, h
    
    If si = 1 Then w = Round(w * arg1): h = Round(h * arg1)
    If si = 2 Then w = arg1: h = arg2
    
    If Scales <> 1 Then
        w = w * Scales
        h = h * Scales
    End If
    
    Dim gC As clsid, gL As Long, index As Long, temp() As String, CrashPath As Long
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
    GdipImageGetFrameCount I, gC, gL
    
    If gL > 0 Then  '动态GIF支持
        temp = Split(path, "\")
        Add 0, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
        With Mems(UBound(Mems))
            .GIF.Count = gL
            ReDim .GIF.frames(gL)
            For index = 0 To gL
                GdipImageSelectActiveFrame I, gC, index
                DC = CreateCDC(w, h)
                PoolCreateFromHdc DC, g
                GdipDrawImageRect g, I, 0, 0, w, h
                PoolDeleteGraphics g
                .GIF.frames(index) = DC
            Next
        End With
        
        Mems(UBound(Mems)).ImgHwnd = I
        Exit Sub
    End If
    
    temp = Split(path, "\")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Name = IIf(Name = "", temp(UBound(temp)), Name)
        .Folder = temp(UBound(temp) - 1)
        .w = w
        .h = h
        .ImgHwnd = I
    End With
    
    For S = 0 To 3
        If NoLoadDirection And S > 0 Then Exit For
        DC = CreateCDC(w, h)
        PoolCreateFromHdc DC, g
        GdipDrawImageRect g, I, IIf(S = 1 Or S = 3, w, 0), IIf(S = 2 Or S = 3, h, 0), w * IIf(S = 1 Or S = 3, -1, 1), h * IIf(S = 2 Or S = 3, -1, 1)
        PoolDeleteGraphics g
        Mems(UBound(Mems)).Imgs(S) = DC
    Next
    
    If w > GW * 0.5 Or h > GH * 0.5 Or NoLoadCrash Then Exit Sub
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim iW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(w + 1, h + 1)
    
    DC = Mems(UBound(Mems)).Imgs(0)
    GdipGetImageWidth I, iW: GdipGetImageHeight I, iH
    For X = 0 To w - 1
        For y = 0 To h - 1
            '取得像素点颜色
            GdipBitmapGetPixel I, X / (w - 1) * iW, y / (h - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4

            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(X, y) = True
        Next
        DoEvents
    Next
    
    Mems(UBound(Mems)).CrashIndex = UBound(ColorLists)
End Sub
Public Sub NewImages(path As String, Optional arg1, Optional arg2)

    If HotLoad Then
        ReDim Preserve AssetsFolder(UBound(AssetsFolder) + 1)
        ReDim Preserve FolderParam1(UBound(FolderParam1) + 1)
        ReDim Preserve FolderParam2(UBound(FolderParam2) + 1)
        If Right(path, 1) <> "\" Then path = path & "\"
        AssetsFolder(UBound(AssetsFolder)) = path
        FolderParam1(UBound(FolderParam1)) = arg1
        FolderParam2(UBound(FolderParam2)) = arg2
        Exit Sub
    End If
    
    Dim f As String, C As Long
    If Right(path, 1) <> "\" Then path = path & "\"
    
    Dim index As Integer, Tree As AssetsTree
    Tree.path = path: ReDim Tree.Files(0)
    index = FindAssetsTree(path, arg1, arg2)
    
    f = Dir(path)
    Do While f <> ""
        C = C + 1
        ReDim Preserve Tree.Files(UBound(Tree.Files) + 1)
        f = Dir()
    Loop
    If index = 0 Then
        PreLoadCount = PreLoadCount + C
    Else
        Tree = AssetsTrees(index)
        For I = 1 To UBound(Tree.Files)
            ReDim Preserve Mems(UBound(Mems) + 1)
            Mems(UBound(Mems)) = Tree.Files(I)
            Mems(UBound(Mems)).copyed = True
        Next
        ReLoadCount = ReLoadCount + UBound(Tree.Files)
        Exit Sub
    End If
    
    f = Dir(path)
    I = 1
    Do While f <> ""
        If IsMissing(arg1) Then
            newImage path & f
        ElseIf IsMissing(arg2) Then
            newImage path & f, arg1
        Else
            newImage path & f, arg1, arg2
        End If
        Tree.Files(I) = Mems(UBound(Mems))
        I = I + 1
        f = Dir()
        DoEvents
        LoadedCount = LoadedCount + 1
    Loop
    
    If index = 0 Then AddAssetsTree Tree, arg1, arg2
End Sub

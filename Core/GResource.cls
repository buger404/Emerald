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

    Dim Path As Long, bmpGraph As Long, BMP As Long, BMP2 As Long, tempGraph As Long
    Dim w As Long, h As Long
    
    GdipGetImageWidth Mems(index).ImgHwnd, w
    GdipGetImageHeight Mems(index).ImgHwnd, h

    GdipCreateBitmapFromScan0 w, h, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP2
    GdipCreateBitmapFromScan0 w * 2, h * 2, ByVal 0, PixelFormat32bppARGB, ByVal 0, BMP

    GdipGetImageGraphicsContext BMP2, tempGraph
    GdipGetImageGraphicsContext BMP, bmpGraph
    GdipSetSmoothingMode tempGraph, SmoothingModeAntiAlias
    GdipSetSmoothingMode bmpGraph, SmoothingModeAntiAlias
    
    PoolCreatePath FillModeWinding, Path
    GdipAddPathEllipseI Path, 1, 1, w * 2 - 2, h * 2 - 2
    GdipSetClipPath bmpGraph, Path, CombineModeReplace
    
    GdipDrawImageRectI bmpGraph, Mems(index).ImgHwnd, 0, 0, w * 2, h * 2
    GdipResetClip bmpGraph
    
    GdipSetInterpolationMode tempGraph, InterpolationModeHighQualityBilinear
    
    GdipDrawImageRectI tempGraph, BMP, 0, 0, w - 1, h - 1

    GdipDeletePath Path
    PoolDeleteGraphics bmpGraph
    PoolDisposeImage BMP
    PoolDeleteGraphics tempGraph
    PoolDisposeImage Mems(index).ImgHwnd
    
    Mems(index).ImgHwnd = BMP2
    
    Call ReDraw(index)
End Sub
Private Sub ApplyEffect(n, kind As GdipEffectType)
    Dim index As Integer
    
    If TypeName(n) = "String" Then
        index = GetImage(n)
    Else
        index = n
    End If
    
    effectimg = Mems(index).ImgHwnd
    
    PoolCreateEffect2 kind, Effect
    
    effectindex = index
End Sub
Private Sub ApplyEffectStep2()
    GdipGetImageWidth effectimg, effectw: GdipGetImageHeight effectimg, effecth
    GdipBitmapApplyEffect effectimg, Effect, NewRectL(0, 0, effectw, effecth), 0, 0, 0
    
    PoolDeleteEffect Effect '垃圾处理
    
    Call ReDraw(effectindex)
End Sub
Public Sub ApplyTintEffect(n, amount As Long, hue As Long)
    Dim e As TintParams
    e.amount = amount
    e.hue = hue
    
    ApplyEffect n, Tint
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplySharpenEffect(n, amount As Single, Radius As Single)
    Dim e As SharpenParams
    e.amount = amount
    e.Radius = Radius

    ApplyEffect n, Sharpen
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBlurEffect(n, Radius As Long, edge As Integer)
    Dim e As BlurParams
    e.Radius = Radius
    e.expandEdge = edge

    ApplyEffect n, Blur
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyBrightnessContrastEffect(n, brightnessLevel As Long, contrastLevel As Long)
    Dim e As BrightnessContrastParams
    e.brightnessLevel = brightnessLevel
    e.contrastLevel = contrastLevel

    ApplyEffect n, BrightnessContrast
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorBalanceEffect(n, cyanRed As Long, magentaGreen As Long, yellowBlue As Long)
    Dim e As ColorBalanceParams
    e.cyanRed = cyanRed
    e.magentaGreen = magentaGreen
    e.yellowBlue = YellowGreen

    ApplyEffect n, ColorBalance
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorCurveEffect(n, adjustment As CurveAdjustments, adjustValue As Long, channel As CurveChannel)
    Dim e As ColorCurveParams
    e.adjustment = adjustment
    e.adjustValue = adjustValue
    e.channel = channel

    ApplyEffect n, ColorCurve
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyColorMatrixEffect(n, m() As Single)
    Dim e As ColorMatrix
    For i = 0 To 4
        For s = 0 To 4
            e.m(i, s) = m(i, s)
        Next
    Next
    ApplyEffect n, ColorMatrix
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyHueSaturationLightnessEffect(n, hueLevel As Long, lightnessLevel As Long, saturationLevel As Long)
    Dim e As HueSaturationLightnessParams
    e.hueLevel = hueLevel
    e.lightnessLevel = lightnessLevel
    e.saturationLevel = saturationLevel

    ApplyEffect n, HueSaturationLightness
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Public Sub ApplyLevelsEffect(n, highlight As Long, midtone As Long, Shadow As Long)
    Dim e As LevelsParams
    e.highlight = highlight
    e.midtone = midtone
    e.Shadow = Shadow

    ApplyEffect n, levels
    GdipSetEffectParameters Effect, e, LenB(e)
    ApplyEffectStep2
End Sub
Private Sub ReDraw(i As Integer)
    If Mems(i).GIF.Count > 0 Then Exit Sub
    
    Dim g As Long
    PoolCreateFromHdc Mems(i).Hwnd, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipGraphicsClear g, 0
    GdipDrawImageRect g, Mems(i).ImgHwnd, 0, 0, Mems(i).w, Mems(i).h
    PoolDeleteGraphics g
    
    For s = 0 To 3
        PoolCreateFromHdc Mems(i).Imgs(s), g
        GdipSetSmoothingMode g, SmoothingModeAntiAlias
        GdipGraphicsClear g, 0
        GdipDrawImageRect g, Mems(i).ImgHwnd, IIf(s = 1 Or s = 3, Mems(i).w, 0), IIf(s = 2 Or s = 3, Mems(i).h, 0), Mems(i).w * IIf(s = 1 Or s = 3, -1, 1), Mems(i).h * IIf(s = 2 Or s = 3, -1, 1)
        PoolDeleteGraphics g
    Next
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim iW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(Mems(i).w + 1, Mems(i).h + 1)
    
    DC = Mems(i).Imgs(0)
    GdipGetImageWidth Mems(i).ImgHwnd, iW: GdipGetImageHeight Mems(i).ImgHwnd, iH
    For x = 0 To Mems(i).w - 1
        For y = 0 To Mems(i).h - 1
            '取得像素点颜色
            GdipBitmapGetPixel Mems(i).ImgHwnd, x / (Mems(i).w - 1) * iW, y / (Mems(i).h - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4
            
            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(x, y) = True
        Next
        DoEvents
    Next
    
    Mems(i).CrashIndex = UBound(ColorLists)
End Sub
Public Sub Add(mem As Long, kind As Integer, name As String, Optional w, Optional h, Optional Folder As String = "")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Hwnd = mem
        .kind = kind
        .name = name
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
    For i = 1 To UBound(Mems)
        If Not Mems(i).copyed Then
            Select Case Mems(i).kind
                Case 0 'DC
                    If Mems(i).Hwnd <> 0 Then DeleteObject Mems(i).Hwnd
                    If Mems(i).ImgHwnd <> 0 Then PoolDisposeImage Mems(i).ImgHwnd
                    For s = 0 To 3
                        If Mems(i).Imgs(s) <> 0 Then DeleteObject Mems(i).Imgs(s)
                    Next
                Case 1 'Graphics
                    PoolDeleteGraphics Mems(i).Hwnd
            End Select
            If Mems(i).GIF.Count > 0 Then
                For s = 0 To Mems(i).GIF.Count
                    DeleteObject Mems(i).GIF.frames(s)
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
    ResourceGIFTime = Mems(index).GIF.time
End Property
Public Property Get ResourceName(index As Integer) As Long
    ResourceName = Mems(index).name
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
Public Function GetImage(ByVal name As String) As Integer
    For i = SearchIndex To UBound(Mems)
        If Mems(i).kind = 0 Then
            If Mems(i).name = name Then GetImage = i: Exit For
        End If
    Next
    
    If HotLoad Then
        If GetImage = 0 Then
            For i = 1 To UBound(AssetsFolder)
                If IsExitAFile(AssetsFolder(i) & name) Then
                    newImage AssetsFolder(i) & name, FolderParam1(i), FolderParam2(i)
                    GetImage = UBound(Mems)
                    Exit For
                End If
            Next
        End If
    End If
End Function
Public Sub newImage(Path As String, Optional arg1, Optional arg2, Optional name As String = "")
    'when arg1 is missing ; i.w = i.w
    'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
    'other ; i.w = arg1
    Dim si As Integer, i As Long, w As Long, h As Long, DC As Long, g As Long
    If IsMissing(arg1) Then si = 0
    If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
    If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
    
    PoolCreateBitmapFromFile StrPtr(Path), i
    GdipGetImageWidth i, w: GdipGetImageHeight i, h
    
    If si = 1 Then w = Round(w * arg1): h = Round(h * arg1)
    If si = 2 Then w = arg1: h = arg2
    
    Dim gC As clsid, gL As Long, index As Long, temp() As String, CrashPath As Long
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
    GdipImageGetFrameCount i, gC, gL
    
    If gL > 0 Then  '动态GIF支持
        temp = Split(Path, "\")
        Add 0, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
        With Mems(UBound(Mems))
            .GIF.Count = gL
            ReDim .GIF.frames(gL)
            For index = 0 To gL
                GdipImageSelectActiveFrame i, gC, index
                DC = CreateCDC(w, h)
                PoolCreateFromHdc DC, g
                GdipDrawImageRect g, i, 0, 0, w, h
                PoolDeleteGraphics g
                .GIF.frames(index) = DC
            Next
        End With
        
        Mems(UBound(Mems)).ImgHwnd = i
        Exit Sub
    End If
    
    temp = Split(Path, "\")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .name = IIf(name = "", temp(UBound(temp)), name)
        .Folder = temp(UBound(temp) - 1)
        .w = w
        .h = h
        .ImgHwnd = i
    End With
    
    For s = 0 To 3
        If NoLoadDirection And s > 0 Then Exit For
        DC = CreateCDC(w, h)
        PoolCreateFromHdc DC, g
        GdipDrawImageRect g, i, IIf(s = 1 Or s = 3, w, 0), IIf(s = 2 Or s = 3, h, 0), w * IIf(s = 1 Or s = 3, -1, 1), h * IIf(s = 2 Or s = 3, -1, 1)
        PoolDeleteGraphics g
        Mems(UBound(Mems)).Imgs(s) = DC
    Next
    
    If w > GW * 0.5 Or h > GH * 0.5 Or NoLoadCrash Then Exit Sub
    
    Dim tempCo As Long, tCo(3) As Byte
    Dim iW As Long, iH As Long

    ReDim Preserve ColorLists(UBound(ColorLists) + 1)
    ReDim ColorLists(UBound(ColorLists)).IsAlpha(w + 1, h + 1)
    
    DC = Mems(UBound(Mems)).Imgs(0)
    GdipGetImageWidth i, iW: GdipGetImageHeight i, iH
    For x = 0 To w - 1
        For y = 0 To h - 1
            '取得像素点颜色
            GdipBitmapGetPixel i, x / (w - 1) * iW, y / (h - 1) * iH, tempCo
            CopyMemory tCo(0), tempCo, 4

            If tCo(3) = 255 Then ColorLists(UBound(ColorLists)).IsAlpha(x, y) = True
        Next
        DoEvents
    Next
    
    Mems(UBound(Mems)).CrashIndex = UBound(ColorLists)
End Sub
Public Sub NewImages(Path As String, Optional arg1, Optional arg2)

    If HotLoad Then
        ReDim Preserve AssetsFolder(UBound(AssetsFolder) + 1)
        ReDim Preserve FolderParam1(UBound(FolderParam1) + 1)
        ReDim Preserve FolderParam2(UBound(FolderParam2) + 1)
        If Right(Path, 1) <> "\" Then Path = Path & "\"
        AssetsFolder(UBound(AssetsFolder)) = Path
        FolderParam1(UBound(FolderParam1)) = arg1
        FolderParam2(UBound(FolderParam2)) = arg2
        Exit Sub
    End If
    
    Dim f As String, c As Long
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    Dim index As Integer, Tree As AssetsTree
    Tree.Path = Path: ReDim Tree.files(0)
    index = FindAssetsTree(Path, arg1, arg2)
    
    f = Dir(Path)
    Do While f <> ""
        c = c + 1
        ReDim Preserve Tree.files(UBound(Tree.files) + 1)
        f = Dir()
    Loop
    If index = 0 Then
        PreLoadCount = PreLoadCount + c
    Else
        Tree = AssetsTrees(index)
        For i = 1 To UBound(Tree.files)
            ReDim Preserve Mems(UBound(Mems) + 1)
            Mems(UBound(Mems)) = Tree.files(i)
            Mems(UBound(Mems)).copyed = True
        Next
        ReLoadCount = ReLoadCount + UBound(Tree.files)
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
        Tree.files(i) = Mems(UBound(Mems))
        i = i + 1
        f = Dir()
        DoEvents
        LoadedCount = LoadedCount + 1
    Loop
    
    If index = 0 Then AddAssetsTree Tree, arg1, arg2
End Sub

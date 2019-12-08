VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald 相关代码

Private Type Page
    p As GPage
    Name As String
End Type
Private Type Transform
    Kind As TranslationKind
    Time As Long
    during As Long
End Type
Dim Pages() As Page, Transform As Transform
Dim mPage As String, aPage As Object
Dim OutOfWindow As Boolean
Dim FirstPlay As Boolean
Public CDC As Long, ODC As Long, FDC As Long, GG As Long, pB As Long, Pen As Long
Public LastPage As String
Public UpdatingPage As String, LockPage As String, UPage As GPage
Public FancyDeepth As Single
Dim FMode As Boolean
Dim CurrentPage As Integer
Dim IsLayered As Boolean
Dim ErrorMark As Boolean, ErrorMark2 As Boolean
Dim F11KEYPress As Boolean, PageMode As Boolean
'========================================================
'   Init
    Public Property Get PageCount() As Long
        PageCount = UBound(Pages)
    End Property
    Public Property Get PageName(Index As Integer) As String
        PageName = Pages(Index).Name
    End Property
    Public Property Get PageController(Index As Integer) As GPage
        Set PageController = Pages(Index).p
    End Property
    Public Property Get FancyMode() As Boolean
        FancyMode = FMode
    End Property
    Public Property Let FancyMode(nV As Boolean)
        If nV Then PaintDC CDC, FDC, 0, 0, alpha:=1
        FMode = nV
    End Property
    Private Sub Class_Initialize()
        Set ECore = Me
        ReDim Pages(0)
        CDC = CreateCDC(RGW, RGH)
        ODC = CreateCDC(RGW, RGH)
        FDC = CreateCDC(RGW, RGH)
        PoolCreateFromHdc CDC, GG
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        PoolCreateSolidFill argb(255, 255, 255, 255), pB
        PoolCreatePen1 argb(255, 255, 255, 255), 1, UnitPixel, Pen
        FancyDeepth = 0.3
    End Sub
    Public Sub Layered(Optional WithAero As Boolean = False)
        SetWindowLongA GHwnd, GWL_EXSTYLE, GetWindowLongA(GHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        If WithAero Then BlurWindow GHwnd
        IsLayered = True
    End Sub
    Public Sub Dispose()
        PoolDeleteGraphics GG
        PoolDeleteBrush pB
        PoolDeletePen Pen
        For i = 1 To UBound(Pages)
            Pages(i).p.Dispose
        Next
        'DeleteObject CDC
        'DeleteObject ODC
    End Sub
'========================================================
'   Manager
    Public Sub Add(p As GPage, Name As String)
        ReDim Preserve Pages(UBound(Pages) + 1)
        With Pages(UBound(Pages))
            Set .p = p
            .Name = Name
        End With
    End Sub
    Public Property Get ActivePage() As String
        ActivePage = mPage
    End Property
    Public Property Let ActivePage(nPage As String)
        LastPage = mPage
        For i = 1 To UBound(Pages)
            If mPage = Pages(i).Name Then CallByName Pages(i).p.SrcClass, "Leave", VbMethod: Exit For
        Next
        mPage = nPage
        For i = 1 To UBound(Pages)
            If nPage = Pages(i).Name Then
                CallByName Pages(i).p.SrcClass, "Enter", VbMethod
                Set aPage = Pages(i).p.SrcClass
                Exit For
            End If
        Next
    End Property
    Public Sub Wheel(Direction As Integer, Depth As Single)
        CallByName aPage, "Wheel", VbMethod, Direction, Depth
    End Sub
'========================================================
'   Draw
    Public Sub Display()
        If ErrorMark2 Then Exit Sub

        Dim o As Boolean, t As Long, played As Boolean
        
        t = Timer * 1000
        
        If ECore Is Me Then
            If (Not FirstPlay) Then
                FirstPlay = True: SysPage.OpenTime = GetTickCount: SysPage.Index = 1
                Call NewTransform
            End If
            If (Not SysPage.DoneMark) Then
                SysPage.Page.Update
                SysPage.Page.Display CDC
                played = True
                GoTo skipPage
            End If
        End If

        For i = 1 To UBound(Pages)
            CurrentPage = i
            If ActivePage = Pages(i).Name Or (Pages(i).p.TopPage And (Not ErrorMark)) Then
                UpdatingPage = Pages(i).Name
                Set UPage = Pages(i).p
                Pages(i).p.Update
                Pages(i).p.Display (CDC)
                If Pages(i).Name = "ErrorScreen" Then ErrorMark2 = True
                played = True
                o = (GetActiveWindow = 0)
                If o <> OutOfWindow Then
                    OutOfWindow = o
                    If OutOfWindow Then
                        CallByName Pages(i).p.SrcClass, "Lostfocus", VbMethod
                    Else
                        CallByName Pages(i).p.SrcClass, "Gotfocus", VbMethod
                    End If
                End If
            End If
        Next
        
skipPage:
        If Not played Then
            GdipGraphicsClear GG, argb(255, 0, 0, 0)
            Suggest "页面""" & ActivePage & """不存在。", ClearOnUpdate, 2
        End If
        
        If GetTickCount - Transform.Time <= Transform.during Then Call Translation
        
        Dim Key As Boolean
        Key = IsKeyPress(VK_F11)
        If Key <> F11KEYPress Then
            F11KEYPress = Key
            If Key Then PageMode = Not PageMode: ECore.NewTransform
        End If
        
        If App.LogMode = 0 Then
            Dim SOrIndex As Integer
            SOrIndex = SysPage.Index
            SysPage.Index = 4
            SysPage.Page.Update
            SysPage.Page.Display (CDC)
            SysPage.Index = SOrIndex
            
            If PageMode Then
                SysPage.Index = 5
                SysPage.Page.Update
                SysPage.Page.Display (CDC)
                SysPage.Index = SOrIndex
            End If
            For i = 1 To UBound(SGS)
                If i > UBound(SGS) Then Exit For
                If SGS(i).ClearTime = ClearOnUpdate Then GoTo ClearSGS
                If SGS(i).ClearTime = ClearOnOnce And GetTickCount - SGS(i).Time >= 1000 Then GoTo ClearSGS
                GoTo NoClear
ClearSGS:
                SGS(i) = SGS(UBound(SGS)): ReDim Preserve SGS(UBound(SGS) - 1): i = i - 1
NoClear:
            Next
            
            If GetTickCount - FPSt >= 1000 Then
                FPSt = GetTickCount: FPS = tFPS: tFPS = 0
                If FPS < 20 Then
                    Suggest "FPS过低。", ClearOnOnce, 2
                End If
                FPSct = FPSctt: FPSctt = 0
            End If
        End If
        
        If FMode Then PaintDC CDC, FDC, 0, 0, alpha:=FancyDeepth
        If IsLayered Then
            Dim bs As BLENDFUNCTION, sz As size
            Dim SrcPoint As POINTAPI
            With bs
                .AlphaFormat = AC_SRC_ALPHA
                .BlendFlags = 0
                .BlendOp = AC_SRC_OVER
                .SourceConstantAlpha = 255
            End With
            sz.cx = RGW: sz.cy = RGH
            'UpdateLayeredWindow GHwnd, 0, ByVal 0, Sz, CDC, ByVal 0, 0, Bs, ULW_ALPHA
            UpdateLayeredWindow GHwnd, GDC, ByVal 0&, sz, IIf(FMode, FDC, CDC), SrcPoint, 0, bs, &H2
        Else
            BitBlt GDC, 0, 0, RGW, RGH, IIf(FMode, FDC, CDC), 0, 0, vbSrcCopy
        End If
        tFPS = tFPS + 1: FPSctt = FPSctt + (Timer * 1000 - t)
        
        If Mouse.state = 2 Then Mouse.state = 0
        
    End Sub
'========================================================
'   Transform
    Public Sub NewTransform(Optional Kind As TranslationKind = transFadeIn, Optional during As Long = 300, Optional newPage)
        PaintDC CDC, ODC
        With Transform
            .during = during
            .Kind = Kind
            .Time = GetTickCount
        End With
        
        If Not IsMissing(newPage) Then
            ActivePage = newPage
        End If
    End Sub
    Private Sub Translation()
        Dim pro As Single
        Select Case Transform.Kind
            Case 0 'FadeIn
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, alpha:=1 - pro
            Case 1 'FadeOut
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 0, 0, 1)
                PaintDC ODC, CDC, alpha:=1 - pro
            Case 2 'ToRight
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * RGW), alpha:=1
            Case 3 'ToLeft
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * RGW), alpha:=1
            Case 4 'ToUp
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * RGW), alpha:=1
            Case 5 'ToDown
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * RGW), alpha:=1
            Case 6 'ToRight(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * RGW), alpha:=1 - pro
            Case 7 'ToLeft(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * RGW), alpha:=1 - pro
            Case 8 'ToUp(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * RGW), alpha:=1 - pro
            Case 9 'ToDown(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * RGW), alpha:=1 - pro
            Case 10 'HighLight
                pro = (GetTickCount - Transform.Time) / Transform.during
                Select Case pro
                    Case Is < 0.2
                        pro = Cubic(pro / 0.2, 0, 1, 1, 1)
                        GdipSetSolidFillColor pB, argb(pro * 255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, RGW + 1, RGH + 1
                    Case Is < 0.8
                        GdipSetSolidFillColor pB, argb(255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, RGW + 1, RGH + 1
                    Case Is < 1
                        pro = Cubic((pro - 0.8) / 0.2, 0, 0, 0, 1)
                        GdipSetSolidFillColor pB, argb((1 - pro) * 255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, RGW + 1, RGH + 1
                End Select
            Case 11 'Dark
                pro = (GetTickCount - Transform.Time) / Transform.during
                Select Case pro
                    Case Is < 0.2
                        pro = Cubic(pro / 0.2, 0, 1, 1, 1)
                        GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, RGW + 1, RGH + 1
                    Case Is < 0.8
                        GdipSetSolidFillColor pB, argb(255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, RGW + 1, RGH + 1
                    Case Is < 1
                        pro = Cubic((pro - 0.8) / 0.2, 0, 0, 0, 1)
                        GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, RGW + 1, RGH + 1
                End Select
            Case 12
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                BitBlt CDC, GW - (pro * RGW), 0, RGW, RGH, CDC, 0, 0, vbSrcCopy
                GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, GW - (pro * RGW), 0, RGW + 1, RGH + 1
                PaintDC ODC, CDC, -(pro * RGW), 0, alpha:=1
                GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, -(pro * RGW), 0, RGW + 1, RGH + 1
            Case 13
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                BitBlt CDC, -RGW + (pro * RGW), 0, RGW, RGH, CDC, 0, 0, vbSrcCopy
                GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, -RGW + (pro * RGW), 0, RGW + 1, RGH + 1
                PaintDC ODC, CDC, (pro * RGW), 0, alpha:=1
                GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, (pro * RGW), 0, RGW + 1, RGH + 1
        End Select
    End Sub
'========================================================

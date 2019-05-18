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
    name As String
End Type
Private Type Transform
    kind As TranslationKind
    time As Long
    during As Long
End Type
Dim Pages() As Page, Transform As Transform
Dim mPage As String, aPage As Object
Dim OutOfWindow As Boolean
Dim FirstPlay As Boolean
Public CDC As Long, ODC As Long, GG As Long, pB As Long
Dim IsLayered As Boolean
'========================================================
'   Init
    Private Sub Class_Initialize()
        Set ECore = Me
        ReDim Pages(0)
        CDC = CreateCDC(GW, GH)
        ODC = CreateCDC(GW, GH)
        GdipCreateFromHDC CDC, GG
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        GdipCreateSolidFill argb(255, 255, 255, 255), pB
    End Sub
    Public Sub Layered(Optional WithAero As Boolean = False)
        SetWindowLongA GHwnd, GWL_EXSTYLE, GetWindowLongA(GHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        If WithAero Then BlurWindow GHwnd
        IsLayered = True
    End Sub
    Public Sub Dispose()
        GdipDeleteGraphics GG
        GdipDeleteBrush pB
        For i = 1 To UBound(Pages)
            Pages(i).p.Dispose
        Next
        DeleteObject CDC
        DeleteObject ODC
    End Sub
'========================================================
'   Manager
    Public Sub Add(p As GPage, name As String)
        ReDim Preserve Pages(UBound(Pages) + 1)
        With Pages(UBound(Pages))
            Set .p = p
            .name = name
        End With
    End Sub
    Public Property Get ActivePage() As String
        ActivePage = mPage
    End Property
    Public Property Let ActivePage(nPage As String)
        For i = 1 To UBound(Pages)
            If mPage = Pages(i).name Then CallByName Pages(i).p.SrcClass, "Leave", VbMethod: Exit For
        Next
        mPage = nPage
        For i = 1 To UBound(Pages)
            If nPage = Pages(i).name Then
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
        Dim o As Boolean, t As Long, played As Boolean
        
        t = Timer * 1000
        
        If ECore Is Me Then
            If (Not FirstPlay) Then
                FirstPlay = True: SysPage.OpenTime = GetTickCount: SysPage.index = 1
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
            If ActivePage = Pages(i).name Then
                Pages(i).p.Update: Pages(i).p.Display (CDC)
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
                Exit For
            End If
        Next
        
skipPage:
        If Not played Then
            GdipGraphicsClear GG, argb(255, 255, 255, 255)
            EF.Writes "页面""" & ActivePage & """不存在", GW / 2, GH / 2, GG, argb(255, 27, 27, 27), 20, 0, 0, StringAlignmentCenter, FontStyleBold
        End If
        
        If GetTickCount - Transform.time <= Transform.during Then Call Translation
        
        If IsLayered Then
            Dim Bs As BLENDFUNCTION, Sz As size
            'Dim SrcPoint As POINTAPI
            With Bs
                .AlphaFormat = AC_SRC_ALPHA
                .BlendFlags = 0
                .BlendOp = AC_SRC_OVER
                .SourceConstantAlpha = 255
            End With
            Sz.cx = GW: Sz.cy = GH
            'UpdateLayeredWindow GHwnd, 0, ByVal 0, Sz, CDC, ByVal 0, 0, Bs, ULW_ALPHA
            UpdateLayeredWindow GHwnd, GDC, ByVal 0&, Sz, CDC, 0, 0, Bs, &H2
        Else
            BitBlt GDC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
        End If
        tFPS = tFPS + 1: FPSctt = FPSctt + (Timer * 1000 - t)
        
        If GetTickCount - FPSt >= 1000 Then
            FPSt = GetTickCount: FPS = tFPS: tFPS = 0
            If FPS < 20 Then
                If FPSWarn < 10 Then FPSWarn = FPSWarn + 1
                If FPSWarn = 10 Then Suggest "FPS过低了噢！你是想要让玩家玩PPT嘛！": FPSWarn = FPSWarn + 1
            ElseIf FPSWarn <> 0 Then
                FPSWarn = 0
            End If
            FPSct = FPSctt: FPSctt = 0
        End If
        
        If Mouse.state = 2 Then Mouse.state = 0
    End Sub
'========================================================
'   Transform
    Public Sub NewTransform(Optional kind As TranslationKind = transFadeIn, Optional during As Long = 300, Optional newPage)
        PaintDC CDC, ODC
        With Transform
            .during = during
            .kind = kind
            .time = GetTickCount
        End With
        
        If Not IsMissing(newPage) Then
            ActivePage = newPage
        End If
    End Sub
    Private Sub Translation()
        Dim pro As Single
        Select Case Transform.kind
            Case 0 'FadeIn
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, Alpha:=1 - pro
            Case 1 'FadeOut
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 0, 0, 1)
                PaintDC ODC, CDC, Alpha:=1 - pro
            Case 2 'ToRight
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * GW), Alpha:=1
            Case 3 'ToLeft
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), Alpha:=1
            Case 4 'ToUp
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), Alpha:=1
            Case 5 'ToDown
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), Alpha:=1
            Case 6 'ToRight(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * GW), Alpha:=1 - pro
            Case 7 'ToLeft(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), Alpha:=1 - pro
            Case 8 'ToUp(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), Alpha:=1 - pro
            Case 9 'ToDown(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), Alpha:=1 - pro
            Case 10 'HighLight
                pro = (GetTickCount - Transform.time) / Transform.during
                Select Case pro
                    Case Is < 0.2
                        pro = Cubic(pro / 0.2, 0, 1, 1, 1)
                        GdipSetSolidFillColor pB, argb(pro * 255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 0.8
                        GdipSetSolidFillColor pB, argb(255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 1
                        pro = Cubic((pro - 0.8) / 0.2, 0, 0, 0, 1)
                        GdipSetSolidFillColor pB, argb((1 - pro) * 255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                End Select
            Case 11 'Dark
                pro = (GetTickCount - Transform.time) / Transform.during
                Select Case pro
                    Case Is < 0.2
                        pro = Cubic(pro / 0.2, 0, 1, 1, 1)
                        GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 0.8
                        GdipSetSolidFillColor pB, argb(255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 1
                        pro = Cubic((pro - 0.8) / 0.2, 0, 0, 0, 1)
                        GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                End Select
        End Select
    End Sub
'========================================================

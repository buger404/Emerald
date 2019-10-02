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
Public CDC As Long, ODC As Long, GG As Long, pB As Long, pen As Long
Public LastPage As String
Public UpdatingPage As String, LockPage As String, UPage As GPage
Dim CurrentPage As Integer
Dim IsLayered As Boolean
Dim ErrorMark As Boolean, ErrorMark2 As Boolean
'========================================================
'   Init
    Private Sub Class_Initialize()
        Set ECore = Me
        ReDim Pages(0)
        CDC = CreateCDC(GW, GH)
        ODC = CreateCDC(GW, GH)
        PoolCreateFromHdc CDC, GG
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        PoolCreateSolidFill argb(255, 255, 255, 255), pB
        PoolCreatePen1 argb(255, 255, 255, 255), 1, UnitPixel, pen
    End Sub
    Public Sub Layered(Optional WithAero As Boolean = False)
        SetWindowLongA GHwnd, GWL_EXSTYLE, GetWindowLongA(GHwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
        If WithAero Then BlurWindow GHwnd
        IsLayered = True
    End Sub
    Public Sub Dispose()
        PoolDeleteGraphics GG
        PoolDeleteBrush pB
        PoolDeletePen pen
        For I = 1 To UBound(Pages)
            Pages(I).p.Dispose
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
        For I = 1 To UBound(Pages)
            If mPage = Pages(I).Name Then CallByName Pages(I).p.SrcClass, "Leave", VbMethod: Exit For
        Next
        mPage = nPage
        For I = 1 To UBound(Pages)
            If nPage = Pages(I).Name Then
                CallByName Pages(I).p.SrcClass, "Enter", VbMethod
                Set aPage = Pages(I).p.SrcClass
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

        For I = 1 To UBound(Pages)
            CurrentPage = I
            If ActivePage = Pages(I).Name Or (Pages(I).p.TopPage And (Not ErrorMark)) Then
                UpdatingPage = Pages(I).Name
                Set UPage = Pages(I).p
                Pages(I).p.Update
                Pages(I).p.Display (CDC)
                If Pages(I).Name = "ErrorScreen" Then ErrorMark2 = True
                played = True
                o = (GetActiveWindow = 0)
                If o <> OutOfWindow Then
                    OutOfWindow = o
                    If OutOfWindow Then
                        CallByName Pages(I).p.SrcClass, "Lostfocus", VbMethod
                    Else
                        CallByName Pages(I).p.SrcClass, "Gotfocus", VbMethod
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
        
        If App.LogMode = 0 Then
            Dim SOrIndex As Integer
            SOrIndex = SysPage.index
            SysPage.index = 4
            SysPage.Page.Update
            SysPage.Page.Display (CDC)
            SysPage.index = SOrIndex
            For I = 1 To UBound(SGS)
                If I > UBound(SGS) Then Exit For
                If SGS(I).ClearTime = ClearOnUpdate Then GoTo ClearSGS
                If SGS(I).ClearTime = ClearOnOnce And GetTickCount - SGS(I).Time >= 1000 Then GoTo ClearSGS
                GoTo NoClear
ClearSGS:
                SGS(I) = SGS(UBound(SGS)): ReDim Preserve SGS(UBound(SGS) - 1): I = I - 1
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
        
        If IsLayered Then
            Dim bs As BLENDFUNCTION, sz As size
            Dim SrcPoint As POINTAPI
            With bs
                .AlphaFormat = AC_SRC_ALPHA
                .BlendFlags = 0
                .BlendOp = AC_SRC_OVER
                .SourceConstantAlpha = 255
            End With
            sz.cx = GW: sz.cy = GH
            'UpdateLayeredWindow GHwnd, 0, ByVal 0, Sz, CDC, ByVal 0, 0, Bs, ULW_ALPHA
            UpdateLayeredWindow GHwnd, GDC, ByVal 0&, sz, CDC, SrcPoint, 0, bs, &H2
        Else
            BitBlt GDC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
        End If
        tFPS = tFPS + 1: FPSctt = FPSctt + (Timer * 1000 - t)
        
        If Mouse.State = 2 Then Mouse.State = 0
        
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
                PaintDC ODC, CDC, (pro * GW), alpha:=1
            Case 3 'ToLeft
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), alpha:=1
            Case 4 'ToUp
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), alpha:=1
            Case 5 'ToDown
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), alpha:=1
            Case 6 'ToRight(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * GW), alpha:=1 - pro
            Case 7 'ToLeft(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), alpha:=1 - pro
            Case 8 'ToUp(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), alpha:=1 - pro
            Case 9 'ToDown(Fade)
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), alpha:=1 - pro
            Case 10 'HighLight
                pro = (GetTickCount - Transform.Time) / Transform.during
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
                pro = (GetTickCount - Transform.Time) / Transform.during
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
            Case 12
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                BitBlt CDC, GW - (pro * GW), 0, GW, GH, CDC, 0, 0, vbSrcCopy
                GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, GW - (pro * GW), 0, GW, GH
                PaintDC ODC, CDC, -(pro * GW), 0, alpha:=1
                GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, -(pro * GW), 0, GW, GH
            Case 13
                pro = Cubic((GetTickCount - Transform.Time) / Transform.during, 0, 1, 1, 1)
                BitBlt CDC, -GW + (pro * GW), 0, GW, GH, CDC, 0, 0, vbSrcCopy
                GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, -GW + (pro * GW), 0, GW, GH
                PaintDC ODC, CDC, (pro * GW), 0, alpha:=1
                GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, (pro * GW), 0, GW, GH
        End Select
    End Sub
'========================================================

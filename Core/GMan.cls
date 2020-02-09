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
    time As Long
    during As Long
End Type
Dim Pages() As Page, Transform As Transform
Dim mPage As String, aPage As Object
Dim OutOfWindow As Boolean
Dim FirstPlay As Boolean
Public CDC As Long, ODC As Long, FDC As Long, NDC As Long, GG As Long, NGG As Long, pB As Long, Pen As Long
Public LastPage As String
Public UpdatingPage As String, LockPage As String, UPage As GPage
Public FreezeMode As Boolean, FreezeDelay As Long, FreezeResetBegin As Boolean
Public FancyDeepth As Single
Dim FMode As Boolean
Dim CurrentPage As Integer
Public IsLayered As Boolean
Dim ErrorMark As Boolean, ErrorMark2 As Boolean
Dim F11KEYPress As Boolean, PageMode As Boolean
Dim FreezeTime As Long
'=======================================================
'   Init
    Public Function SimpleMsg(ByVal Content As String, ByVal Title As String, buttons() As String, Optional ByVal w As Long = -1, Optional ByVal h As Long = -1, Optional UseBlur As Boolean = True, Optional Radius As Long = 60) As Integer
        MsgType = 1
        MsgW = IIf(w = -1, RGW * 0.6, w): MsgH = IIf(h = -1, RGH * 0.4, h)
        MsgBtnFocus = -1: MsgBtnFocusTime = 0
        MsgContent = Content: MsgTitle = Title
        ReDim MsgButtons(UBound(buttons))
        For I = 0 To UBound(buttons)
            MsgButtons(I) = buttons(I)
        Next
        SysPage.MsgButton = -1
        Dim lState As Boolean
        lState = ECore.FreezeMode
        ECore.FreezeMode = False
        
        MsgEnterTime = GetTickCount
        ECore.NewTransform transFadeIn, 500
        
        If UseBlur Then
            Dim frm As Form
            For Each frm In VB.Forms
                If frm.Hwnd = GHwnd Then Exit For
            Next frm
            BlurTo SysPage.BDC, ECore.CDC, frm, Radius
        Else
            PaintDC ECore.CDC, SysPage.BDC, alpha:=1
        End If
        
        Do While SysPage.MsgButton = -1
            ECore.Display: DoEvents
        Loop
        
        SimpleMsg = SysPage.MsgButton
        ECore.NewTransform transFadeIn, 500
        ECore.FreezeMode = lState
        
    End Function
    Public Function ScreenMsg(ByVal Content As String, Optional UseBlur As Boolean = True, Optional Radius As Long = 60) As Integer
        MsgType = 2
        MsgContent = Content
        SysPage.MsgButton = -1
        Dim lState As Boolean
        lState = ECore.FreezeMode
        ECore.FreezeMode = False
        
        ECore.NewTransform transFadeIn, 500
        
        If UseBlur Then
            Dim frm As Form
            For Each frm In VB.Forms
                If frm.Hwnd = GHwnd Then Exit For
            Next frm
            BlurTo SysPage.BDC, ECore.CDC, frm, Radius
        Else
            PaintDC ECore.CDC, SysPage.BDC, alpha:=1
        End If
        
        MsgEnterTime = GetTickCount
        
        Do While SysPage.MsgButton = -1
            ECore.Display: DoEvents
        Loop
        
        ScreenMsg = SysPage.MsgButton
        ECore.NewTransform transFadeIn, 500
        ECore.FreezeMode = lState
        
    End Function
    Public Function TipMsg(ByVal Content As String, Optional ByVal Delay As Long = 2000, Optional ByVal BackColor As Long = -1, Optional ByVal TextColor As Long = -1) As Integer
        If MsgTipTime = -2 Then Exit Function
        
        If CheckMouse2 >= mMouseIn Then
            If MsgBoundRect.X <> DrawF.X Or MsgBoundRect.y <> DrawF.y Or MsgBoundRect.Width <> DrawF.Width Or MsgBoundRect.Height <> DrawF.Height Then
                MsgBoundRect = DrawF: MsgTipTime = GetTickCount: Exit Function
            Else
                If IsKeyPress(VK_LCONTROL) = True Or IsKeyPress(VK_RCONTROL) = True Then MsgTipTime = GetTickCount - Delay
                If GetTickCount - MsgTipTime < Delay Then Exit Function
                If GetTickCount - MsgTipTime >= Delay Then MsgTipTime = -1
            End If
        Else
            Exit Function
        End If
        Dim w As Long, h As Long
        w = EF.GetWidth(GG, Content, 18, StringAlignmentNear, FontStyleBold)
        h = EF.GetHeight(GG, Content, 18, StringAlignmentNear, FontStyleBold)
        MsgW = w + 20 * 2: MsgH = h + 5 * 2
        MsgX = DrawF.X + DrawF.Width / 2 - MsgW / 2: MsgY = DrawF.y - MsgH - 10
        MsgTipDirection = 0
        If MsgY < 0 Then MsgY = DrawF.y + DrawF.Height: MsgTipDirection = 1
        If MsgX < 0 Then MsgX = 0
        If MsgX > GW - MsgW Then MsgX = GW - MsgW
        MsgTipColor = 1
        PaintDC2 CDC, NDC, 0, 0, 3, 3, MsgX, MsgY, MsgW, MsgH, 1
        Dim temp(3) As Byte, Color As Long, total As Long
        Color = GetPixel(NDC, 1, 1)
        CopyMemory temp(0), Color, 4
        For I = 0 To 3
            total = total + temp(I)
        Next
        If total / 3 <= 128 Then MsgTipColor = 0
        
        MsgType = 3: MsgContent = Content
        SysPage.MsgButton = -1
        Dim lState As Boolean
        lState = ECore.FreezeMode
        ECore.FreezeMode = False
        
        Dim frm As Form
        For Each frm In VB.Forms
            If frm.Hwnd = GHwnd Then Exit For
        Next frm
        BlurTo SysPage.BDC, ECore.CDC, frm, 255
        
        MsgBkColor = BackColor
        If MsgBkColor = -1 Then MsgBkColor = IIf(MsgTipColor = 0, argb(180, 255, 255, 255), argb(180, 0, 0, 0))
        MsgFgColor = TextColor
        If MsgFgColor = -1 Then MsgFgColor = IIf(MsgTipColor = 0, argb(255, 0, 0, 0), argb(255, 255, 255, 255))
        
        ECore.FreezeMode = lState
        ECore.FreezeReset
        
        MsgBtnFocusTime = GetTickCount: MsgBtnFocus = -1
        
        If MsgTipTime = -1 Then MsgTipTime = -2
    End Function
    Public Sub FreezeReset()
        FreezeTime = GetTickCount
    End Sub
    Public Property Get FreezeAvailable() As Boolean
        If Not FreezeMode Then Exit Property
        Dim Ret As Boolean
        Ret = False
        If FreezeMode Then Ret = (GetTickCount - FreezeTime >= FreezeDelay)
        FreezeAvailable = Ret
    End Property
    Public Property Get PageCount() As Long
        PageCount = UBound(Pages)
    End Property
    Public Property Get PageName(index As Integer) As String
        PageName = Pages(index).Name
    End Property
    Public Property Get PageController(index As Integer) As GPage
        Set PageController = Pages(index).p
    End Property
    Public Property Get FancyMode() As Boolean
        FancyMode = FMode
    End Property
    Public Property Let FancyMode(nV As Boolean)
        If nV Then PaintDC CDC, FDC, 0, 0, alpha:=1
        FMode = nV
    End Property
    Private Sub Class_Initialize()
        If Not (ECore Is Nothing) Then
            Suggest "已经有一个页面管理器在运作了。", NeverClear, 2
            'Set Me = Nothing
            Exit Sub
        End If
        Set ECore = Me
        ReDim Pages(0)
        CDC = CreateCDC(RGW, RGH)
        ODC = CreateCDC(RGW, RGH)
        FDC = CreateCDC(RGW, RGH)
        NDC = CreateCDC(RGW, RGH)
        PoolCreateFromHdc CDC, GG
        PoolCreateFromHdc NDC, NGG
        GdipSetSmoothingMode GG, SmoothingModeAntiAlias
        GdipSetTextRenderingHint GG, TextRenderingHintAntiAlias
        PoolCreateSolidFill argb(255, 255, 255, 255), pB
        PoolCreatePen1 argb(255, 255, 255, 255), 1, UnitPixel, Pen
        FancyDeepth = 0.3
        FreezeDelay = 3000
        Call FreezeReset
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
        
        If Not HasCheckUpdate Then
            If App.LogMode = 0 Then Call CheckUpdate
        End If
        
        If ECore Is Me Then
            If (Not FirstPlay) Then
                FirstPlay = True: SysPage.OpenTime = GetTickCount: SysPage.index = 1
                Call NewTransform
            End If
            If (Not SysPage.DoneMark) Then
                SysPage.Page.Update
                SysPage.Page.Display CDC
                If SysPage.MsgButton = -1 Then GoTo DrawMessage
                played = True: GoTo skipPage
            End If
        End If

        ChoosePosition = False
        Dim OrIndex As Integer, NDCMark As Boolean
        If IsKeyPress(VK_DECIMAL) And (App.LogMode = 0) Then
            OrIndex = SysPage.index
            SysPage.index = 6
            SysPage.Page.Update
            GdipGraphicsClear NGG, 0
            SysPage.Page.Display NDC
            SysPage.index = OrIndex
            NDCMark = True
        End If
            
        If SysPage.MsgButton = -1 Then
DrawMessage:
            OrIndex = SysPage.index
            SysPage.index = 7
            SysPage.Page.Update
            If MsgType = 3 Then GdipGraphicsClear NGG, 0
            SysPage.Page.Display IIf(MsgType = 3, NDC, CDC)
            SysPage.index = OrIndex
            If MsgType = 3 Then
                NDCMark = True
            Else
                played = True: GoTo skipPage
            End If
        End If
        
        If SysPage.DoneMark = False Then played = True: GoTo skipPage
        
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
        
        If NDCMark Then
            PaintDC NDC, CDC, 0, 0, alpha:=1
        End If
        
        If GetTickCount - Transform.time <= Transform.during Then Call Translation
        
        Dim Key As Boolean
        Key = IsKeyPress(VK_F11)
        If Key <> F11KEYPress Then
            F11KEYPress = Key
            If Key Then PageMode = Not PageMode: ECore.NewTransform
        End If
        
        If App.LogMode = 0 Then
            Dim SOrIndex As Integer
            SOrIndex = SysPage.index
            SysPage.index = 4
            SysPage.Page.Update
            SysPage.Page.Display (CDC)
            SysPage.index = SOrIndex
            
            If PageMode Then
                SysPage.index = 5
                SysPage.Page.Update
                SysPage.Page.Display (CDC)
                SysPage.index = SOrIndex
            End If
            For I = 1 To UBound(SGS)
                If I > UBound(SGS) Then Exit For
                If SGS(I).ClearTime = ClearOnUpdate Then GoTo ClearSGS
                If SGS(I).ClearTime = ClearOnOnce And GetTickCount - SGS(I).time >= 1000 Then GoTo ClearSGS
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
                For I = 0 To 19
                    FPSRecord(I) = FPSRecord(I + 1)
                Next
                FPSRecord(20) = FPS
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
        
        If Mouse.State = 2 Then Mouse.State = 0
        
        If FreezeResetBegin Then FreezeResetBegin = False: Call FreezeReset
    End Sub
'========================================================
'   Transform
    Public Sub NewTransform(Optional Kind As TranslationKind = transFadeIn, Optional during As Long = 300, Optional newPage)
        PaintDC CDC, ODC
        With Transform
            .during = during
            .Kind = Kind
            .time = GetTickCount
        End With
        
        If Not IsMissing(newPage) Then
            ActivePage = newPage
        End If
    End Sub
    Private Sub Translation()
        Dim pro As Single
        Select Case Transform.Kind
            Case 0 'FadeIn
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, alpha:=1 - pro
            Case 1 'FadeOut
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 0, 0, 1)
                PaintDC ODC, CDC, alpha:=1 - pro
            Case 2 'ToRight
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * RGW), alpha:=1
            Case 3 'ToLeft
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * RGW), alpha:=1
            Case 4 'ToUp
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * RGW), alpha:=1
            Case 5 'ToDown
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * RGW), alpha:=1
            Case 6 'ToRight(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * RGW), alpha:=1 - pro
            Case 7 'ToLeft(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * RGW), alpha:=1 - pro
            Case 8 'ToUp(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * RGW), alpha:=1 - pro
            Case 9 'ToDown(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * RGW), alpha:=1 - pro
            Case 10 'HighLight
                pro = (GetTickCount - Transform.time) / Transform.during
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
                pro = (GetTickCount - Transform.time) / Transform.during
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
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                BitBlt CDC, GW - (pro * RGW), 0, RGW, RGH, CDC, 0, 0, vbSrcCopy
                GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, GW - (pro * RGW), 0, RGW + 1, RGH + 1
                PaintDC ODC, CDC, -(pro * RGW), 0, alpha:=1
                GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, -(pro * RGW), 0, RGW + 1, RGH + 1
            Case 13
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                BitBlt CDC, -RGW + (pro * RGW), 0, RGW, RGH, CDC, 0, 0, vbSrcCopy
                GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, -RGW + (pro * RGW), 0, RGW + 1, RGH + 1
                PaintDC ODC, CDC, (pro * RGW), 0, alpha:=1
                GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                GdipFillRectangle GG, pB, (pro * RGW), 0, RGW + 1, RGH + 1
        End Select
    End Sub
'========================================================

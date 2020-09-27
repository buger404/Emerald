Attribute VB_Name = "GCore"
'   Emerald 相关代码
'========================================================
'   DPI适应
    Public Declare Function SetProcessDpiAwareness Lib "SHCORE.DLL" (ByVal DPImodel As Long) As Long
'=========================================================================
    Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
    Public Type MState
        State As Integer
        button As Integer
        X As Single
        y As Single
    End Type
    Public Enum PlayStateMark
        musStopped = 0
        musPlaying = 1
        musStalled = 2
        musPaused = 3
    End Enum
    Public Enum imgIndex
        imgGetWidth = 0
        imgGetHeight = 1
        imgGetGIFFrameCount = 2
    End Enum
    Public Enum MButtonState
        mMouseOut = 0
        mMouseIn = 1
        mMouseDown = 2
        mMouseUp = 3
    End Enum
    Public Enum PosAlign
        posNormal = 0
        posOnCenter = 1
        posOnLeft = 4
        posOnTop = 5
        posOnRight = 2
        posOnBottom = 3
    End Enum
    Public Enum TranslationKind
        transFadeIn = 0
        transFadeOut = 1
        transToRight = 2
        transToLeft = 3
        transToUp = 4
        transToDown = 5
        transToRightFade = 6
        transToLeftFade = 7
        transToUpFade = 8
        transToDownFade = 9
        transHighLight = 10
        transFallDark = 11
        transDarkTo = 12
        transDarkReturn = 13
    End Enum
    Public Type GGIF
        time As Long
        frames() As Long
        tick As Long
        Count As Long
    End Type
    Public Type GMem
        GIF As GGIF
        Kind As Integer
        Hwnd As Long
        ImgHwnd As Long
        Imgs(3) As Long
        Name As String
        Folder As String
        w As Long
        h As Long
        copyed As Boolean
        CrashIndex As Long
    End Type
    Public Type AssetsTree
        Files() As GMem
        path As String
        arg1 As Variant
        arg2 As Variant
    End Type
    Public Enum ImgDirection
        DirNormal = 0
        DirHorizontal = 1
        DirVertical = 2
        DirHorizontalVertical = 3
    End Enum
    Public Type GraphicsBound
        X As Single
        y As Single
        Width As Single
        Height As Single
        WSc As Single
        HSc As Single
        CrashIndex As Long
        Shape As Long
        Strings As String
    End Type
    Public Type ColorCollection
        IsAlpha() As Boolean
    End Type
    Public Enum SuggestClearTime
        NeverClear = 0
        ClearOnUpdate = 1
        ClearOnOnce = 2
    End Enum
    Public Type Suggestion
        Content As String
        Deepth As Long
        time As Long
        ClearTime As SuggestClearTime
    End Type
    Public Enum EScalePolicy
        ScaleDefault = 0
        ScaleFullScreen = 1
        ScaleSuitable = 2
    End Enum
    Public Enum ERenderMode
        RenderModeDefault = 0
        RenderModeFillPath = 1
        RenderModeClearType = 2
    End Enum
    Public Type ScrollArea
        Width As Long
        Height As Long
        Graphics As Long
        DC As Long
    End Type
    Public SGS() As Suggestion, SGTime As Long
    Public ColorLists() As ColorCollection
    Public ECore As GMan, EF As GFont, EAni As Object, ESave As GSaving, EMusic As GMusicList
    Public GHwnd As Long, GDC As Long, GW As Single, GH As Single, RGW As Single, RGH As Single
    Public Mouse As MState, DrawF As GraphicsBound
    Public FPS As Long, FPSt As Long, tFPS As Long, FPSct As Long, FPSctt As Long
    Public SysPage As GSysPage
    Public PreLoadCount As Long, LoadedCount As Long, ReLoadCount As Long
    Public FPSWarn As Long
    Public EmeraldInstalled As Boolean
    Public BassInstalled As Boolean
    '版本号说明：[年份][月份][日期][版次]
    Public Const Version As Long = 20092701     '20年9月27日，第1个版本
    Public TextHandle As Long, WaitChr As String, LastUpdateTime As Long, HighCPUPermission As Integer
    Public XPMode As Boolean
    Public Scales As Single
    Public FullScreenMark As Boolean
    Public UsePaint As Boolean
    Public MsgButtons() As String, MsgTitle As String, MsgContent As String, MsgW As Long, MsgH As Long, MsgX As Long, MsgY As Long
    Public MsgBtnFocus As Integer, MsgBtnFocusTime As Long, MsgType As Integer, MsgTipDirection As Integer, MsgTipColor As Integer
    Public MsgBoundRect As GraphicsBound, MsgTipTime As Long, MsgBkColor As Long, MsgFgColor As Long, MsgEnterTime As Long
    
    Public AssetsTrees() As AssetsTree
    Dim LastKeyUpRet As Boolean
    Dim Wndproc As Long
'================================================================================
    '读取INI
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
    Dim FSO As Object
    Public Function IsExitAFile(filespec As String) As Boolean
        If FSO Is Nothing Then Set FSO = PoolCreateObject("Scripting.FileSystemObject")
        
        IsExitAFile = FSO.fileExists(filespec)
    End Function
'================================================================================
'   运行时
'   读取INI文件
'   <SectionName:标题名称,KeyName:项名称,IniFileName:INI文件路径>
    Private Function ReadINI(ByVal SectionName As String, ByVal KeyName As String, ByVal IniFileName As String) As String
        Dim strBuf As String
        strBuf = String(128, 0)
        GetPrivateProfileString StrPtr(SectionName), StrPtr(KeyName), StrPtr(""), StrPtr(strBuf), 128, StrPtr(IniFileName)
        strBuf = Left(strBuf, InStr(strBuf, Chr(0)))
        ReadINI = strBuf
    End Function
    Public Sub OutPutDebug(Str As String)
        Open App.path & "\debug.txt" For Append As #1
        Print #1, Now & "    " & Str
        Close #1
    End Sub
'================================================================================
'   Init
    Public Function StrArray(ParamArray Members()) As String()
        Dim I As Integer, R() As String
        ReDim R(UBound(Members))
        For I = 0 To UBound(Members)
            R(I) = Members(I)
        Next
        StrArray = R
    End Function
    Public Sub ShowSuggestion()
        If UBound(SGS) = 0 Then
            MsgBox "没有任何建议。", 48, "Emerald 建议中心"
        Else
            Dim Ret As String, sRet As String
            Ret = "共 " & UBound(SGS) & " 项建议。" & vbCrLf & vbCrLf
            For I = 1 To UBound(SGS)
                Select Case SGS(I).Deepth
                    Case 0: sRet = "(无关紧要)"
                    Case 1: sRet = "(需要留意)"
                    Case 2: sRet = "(*紧急)"
                    Case 3: sRet = "(*非常紧急*)"
                End Select
                Ret = Ret & sRet & "    " & SGS(I).Content & vbCrLf
            Next
            MsgBox Ret, 48, "Emerald 建议中心"
        End If
    End Sub
    Public Sub SaveSettings(Data As GSaving)
        Data.PutData "DebugMode", DebugMode
        Data.PutData "DisableLOGO", DisableLOGO
        Data.PutData "HideLOGO", HideLOGO
    End Sub
    Public Function CreateScrollArea(ByVal Width As Long, ByVal Height As Long) As ScrollArea
        With CreateScrollArea
            .DC = CreateCDC(Width, Height)
            PoolCreateFromHdc .DC, .Graphics
            .Width = Width
            .Height = Height
            GdipSetSmoothingMode .Graphics, SmoothingModeAntiAlias
            GdipSetTextRenderingHint .Graphics, TextRenderingHintAntiAlias
        End With
    End Function
    Public Sub StartScrollArea(Page As GPage, Area As ScrollArea)
        If Page.ScrollMode Then Suggest "请先结束上一个卷轴区域。", ClearOnUpdate, 1: Exit Sub
        Page.OODC = Page.CDC: Page.OOGG = Page.GG
        Page.CDC = Area.DC: Page.GG = Area.Graphics
        Page.ScrollWidth = Area.Width: Page.ScrollHeight = Area.Height
        Page.ScrollMode = True
    End Sub
    Public Sub EndScrollArea(Page As GPage, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal alpha As Single = 1)
        If Not Page.ScrollMode Then Suggest "请先启动一个卷轴区域。", ClearOnUpdate, 1: Exit Sub
        If Width = -1 Then Width = Page.ScrollWidth
        If Height = -1 Then Height = Page.ScrollHeight
        PaintDC Page.CDC, Page.OODC, X, y, cx, cy, Width, Height, alpha
        Page.CDC = Page.OODC: Page.GG = Page.OOGG
        Page.ScrollMode = False
    End Sub
    Public Sub GetSettings(Optional SkipDebug As Boolean = False)
        If App.LogMode <> 0 And SkipDebug = False Then Exit Sub
    
        Dim Data As New GSaving
        Data.Create "Emerald.Core"
        Data.AutoSave = True
        
        If Data.GetData("DebugMode") = "" Then
            Call SaveSettings(Data)
        End If
        
        DebugSwitch.DebugMode = Val(Data.GetData("DebugMode"))
        DebugSwitch.DisableLOGO = Val(Data.GetData("DisableLOGO"))
        DebugSwitch.HideLOGO = Val(Data.GetData("HideLOGO"))
        DebugSwitch.HideSuggest = Val(Data.GetData("HideSuggest"))
        
        Set Data = Nothing
    End Sub
    Public Sub ScaleGame(ByVal Percent As Single, ScaleMode As EScalePolicy)
        Select Case ScaleMode
            Case EScalePolicy.ScaleDefault
                RGW = GW * Percent: RGH = GH * Percent: Scales = Percent
            Case EScalePolicy.ScaleFullScreen
                Scales = (Screen.Height / Screen.TwipsPerPixelY) / GH
                RGW = Screen.Width / Screen.TwipsPerPixelX: RGH = Screen.Height / Screen.TwipsPerPixelY
            Case EScalePolicy.ScaleSuitable
                Scales = ((Screen.Width / Screen.TwipsPerPixelX) * Percent) / GW
                RGW = (Screen.Width / Screen.TwipsPerPixelX) * Percent: RGH = (Screen.Height / Screen.TwipsPerPixelY) * Percent
        End Select

        Dim DPI As Long
        DPI = 1440 / Screen.TwipsPerPixelX
        If (GetWindowLongA(GHwnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION Then
            SetWindowPos GHwnd, 0, 0, 0, RGW + 3 * Int(DPI / 96), RGH + 26 * Int(DPI / 96), SWP_NOMOVE Or SWP_NOZORDER
        Else
            SetWindowPos GHwnd, 0, 0, 0, RGW - 2 * Int(DPI / 96), RGH - 2 * Int(DPI / 96), SWP_NOMOVE Or SWP_NOZORDER
        End If
        ReleaseDC GHwnd, GDC
        GDC = GetDC(GHwnd)
    End Sub
    Public Sub StartEmerald(Hwnd As Long, w As Long, h As Long, Optional DPIPolicy As Boolean = True)
        If w >= Screen.Width / Screen.TwipsPerPixelX * 0.9 And h >= Screen.Height / Screen.TwipsPerPixelY * 0.9 Then
            Dim Win As EmeraldWindow, FSPer As Integer
            Set Win = New EmeraldWindow
            FSPer = Win.NewPermissionDialog("全屏", "该应用请求覆盖你的屏幕" & vbCrLf & "为了最佳应用体验效果可能需要全屏" & vbCrLf & "拒绝后将使用默认窗口大小" & vbCrLf & "请求授权的应用：" & App.Title)
            If FSPer = 0 Then w = 800: h = 600
        End If

        HighCPUPermission = -1
    
        ReDim ChooseLines(0)
    
        Scales = 1
    
        ReDim ColorLists(0)
        ReDim SGS(0)
            
        Call InitPool
        
        Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each objItem In colItems
            strOSversion = objItem.Version
        Next
    
        Select Case Val(Split(strOSversion, ".")(0))
        Case Is <= "5"
            XPMode = True
        End Select
    
        Call GetSettings
    
        If DebugMode Then
            If App.LogMode <> 0 Then MsgBox "错误：生成时未关闭Debug模式。": End
        End If
        
        ReDim AssetsTrees(0)
        
        InitGDIPlus
        
        GHwnd = Hwnd: GW = w: GH = h
        RGW = GW: RGH = GH
        Dim DPI As Long
        DPI = 1440 / Screen.TwipsPerPixelX
        If (GetWindowLongA(Hwnd, GWL_STYLE) And WS_CAPTION) = WS_CAPTION Then
            SetWindowPos Hwnd, 0, 0, 0, w + 3 * Int(DPI / 96), h + 26 * Int(DPI / 96), SWP_NOMOVE Or SWP_NOZORDER
        Else
            SetWindowPos Hwnd, 0, 0, 0, w - 2 * Int(DPI / 96), h - 2 * Int(DPI / 96), SWP_NOMOVE Or SWP_NOZORDER
        End If
        
        GDC = GetDC(Hwnd)
        If App.LogMode <> 0 Then Wndproc = SetWindowLongA(Hwnd, GWL_WNDPROC, AddressOf Process)
        
        Set EAni = New GAnimation
        Set SysPage = New GSysPage
        
        If Val(GetWinNTVersion) > 6.1 Then               '如果当前系统版本高于win7
            If DPIPolicy Then
                SetProcessDpiAwareness 2&                '调用API使本程序在高DPI情况下不模糊
            End If
        End If
        
        If DebugMode Then
            Debuginfo.Show
            Debuginfo.Hide
            DebugWindow.Show
        End If
        
        If ReLoadCount > LoadedCount Then Suggest "重复加载的资源数量过多。", NeverClear, 1
        
    End Sub
    Public Sub Suggest(Text As String, Clears As SuggestClearTime, Deepth As Long)
        If DebugSwitch.HideSuggest Then Exit Sub
        ReDim Preserve SGS(UBound(SGS) + 1)
        With SGS(UBound(SGS))
            .Content = Text
            .ClearTime = Clears
            .time = GetTickCount
            .Deepth = Deepth
        End With
        SGTime = GetTickCount
    End Sub
    Public Sub EndEmerald()
        If DebugMode Then
            Unload Debuginfo
            Unload DebugWindow
        End If
        
        If App.LogMode <> 0 Then SetWindowLongA GHwnd, GWL_WNDPROC, Wndproc
        'If Not (ECore Is Nothing) Then ECore.Dispose
        'If Not (EF Is Nothing) Then EF.Dispose
        Call DestroyPool
        
        TerminateGDIPlus
        If BassInstalled Then BASS_Free
    End Sub
    Public Sub MakeFont(ByVal Name As String)
        Set EF = New GFont
        EF.MakeFont Name
    End Sub
'========================================================
'   RunTime
    Public Function ToTime(time) As String
        ToTime = Int(time / 60) & ":" & format(time Mod 60, "00")
    End Function
    Public Function Process(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        On Error GoTo sth

        If uMsg = WM_MOUSEWHEEL Then
            Dim Direction As Integer, Strong As Single
            Direction = IIf(wParam < 0, -1, 1): Strong = Abs(wParam / 7864320)
            ECore.Wheel Direction, Strong
        End If
        
        If ECore.IsLayered Then
            If ((uMsg = 132)) And (UsePaint = True) Then
                ECore.Display
            End If
        Else
            If ((uMsg = WM_PAINT)) And (UsePaint = True) Then
                ECore.Display
            End If
        End If
        
last:
        Process = CallWindowProcA(Wndproc, Hwnd, uMsg, wParam, lParam)
sth:

    End Function
'   取得当前系统的WinNT版本
    Public Function GetWinNTVersion() As String
        Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
        strComputer = "."
        Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
        
        For Each objItem In colItems
            strOSversion = objItem.Version
        Next
        
        GetWinNTVersion = Left(strOSversion, 3)
    End Function
    Public Sub BlurTo(DC As Long, srcDC As Long, buffWin As Form, Optional Radius As Long = 60)
        If XPMode Then BitBlt DC, 0, 0, GW, GH, srcDC, 0, 0, vbSrcCopy: Exit Sub
        
        Dim I As Long, g As Long, e As Long, B As BlurParams, w As Long, h As Long
        '粘贴到缓冲窗口
        buffWin.AutoRedraw = True
        BitBlt buffWin.hdc, 0, 0, GW, GH, srcDC, 0, 0, vbSrcCopy: buffWin.Refresh
        
        '创建Bitmap
        GdipCreateBitmapFromHBITMAP buffWin.Image.handle, buffWin.Image.hpal, I
        
        '模糊操作
        PoolCreateEffect2 GdipEffectType.Blur, e: B.Radius = Radius: GdipSetEffectParameters e, B, LenB(B)
        GdipGetImageWidth I, w: GdipGetImageHeight I, h
        GdipBitmapApplyEffect I, e, NewRectL(0, 0, w, h), 0, 0, 0
        
        '画~
        PoolCreateFromHdc DC, g
        GdipDrawImage g, I, 0, 0
        PoolDisposeImage I: PoolDeleteGraphics g: PoolDeleteEffect e '垃圾处理
        buffWin.AutoRedraw = False
    End Sub
    Public Sub BlurImg(img As Long, Radius As Long)
        If XPMode Then Exit Sub
    
        Dim B As BlurParams, e As Long, w As Long, h As Long
        
        '模糊操作
        
        PoolCreateEffect2 GdipEffectType.Blur, e: B.Radius = Radius: GdipSetEffectParameters e, B, LenB(B)
        GdipGetImageWidth img, w: GdipGetImageHeight img, h
        GdipBitmapApplyEffect img, e, NewRectL(0, 0, w, h), 0, 0, 0
        
        '画~
        PoolDeleteEffect e '垃圾处理
    End Sub
    Public Sub PaintDC2(DC As Long, destDC As Long, X As Long, y As Long, w As Long, h As Long, cx As Long, cy As Long, CW As Long, ch As Long, alpha As Single)
        Dim B As BLENDFUNCTION, index As Integer, bl As Long
        
        If Not IsMissing(alpha) Then
            If alpha < 0 Then alpha = 0
            If alpha > 1 Then alpha = 1
            With B
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(alpha * 255)
            End With
            CopyMemory bl, B, 4
        End If
        
        AlphaBlend destDC, X, y, w, h, DC, cx, cy, CW, ch, bl
    End Sub
    Public Sub PaintDC(DC As Long, destDC As Long, Optional X As Long = 0, Optional y As Long = 0, Optional cx As Long = 0, Optional cy As Long = 0, Optional CW, Optional ch, Optional alpha)
        Dim B As BLENDFUNCTION, index As Integer, bl As Long
        
        If Not IsMissing(alpha) Then
            If alpha < 0 Then alpha = 0
            If alpha > 1 Then alpha = 1
            With B
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(alpha * 255)
            End With
            CopyMemory bl, B, 4
        End If
        
        If IsMissing(CW) Then CW = RGW - cx
        If IsMissing(ch) Then ch = RGH - cy
        
        If IsMissing(alpha) Then
            BitBlt destDC, X, y, CW, ch, DC, cx, cy, vbSrcCopy
        Else
            AlphaBlend destDC, X, y, CW, ch, DC, cx, cy, CW, ch, bl
        End If
    End Sub
    Function Cubic(t As Single, arg0 As Single, arg1 As Single, arg2 As Single, arg3 As Single) As Single
        'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
        'Attention:all the args must in this area (0~1)
        'Cubic = HtmlCubic(t, arg0, arg1, arg2, arg3)
        Cubic = (arg0 * ((1 - t) ^ 3)) + (3 * arg1 * t * ((1 - t) ^ 2)) + (3 * arg2 * (t ^ 2) * (1 - t)) + (arg3 * (t ^ 3))
    End Function
'========================================================
'   Mouse
    Public Sub UpdateMouse(X As Single, y As Single, State As Long, button As Integer)
        With Mouse
            .X = X
            .y = y
            .State = State
            .button = button
        End With
    End Sub
    Public Function CheckMouse(ByVal X As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long) As MButtonState
        'Return Value:0=none,1=in,2=down,3=up
        If Scales <> 1 Then
            X = X * Scales: y = y * Scales
            w = w * Scales
            h = h * Scales
        End If
        If Debug_mouse Then
            GdipSetSolidFillColor ECore.pB, argb(20, 255, 0, 0)
            GdipFillRectangle ECore.UPage.GG, ECore.pB, X, y, w, h
        End If
        If ECore.LockPage <> "" Then
            If ECore.LockPage <> ECore.UpdatingPage Then Exit Function
        End If
        If Mouse.X >= X And Mouse.y >= y And Mouse.X <= X + w And Mouse.y <= y + h Then
            If ECore.FreezeMode Then ECore.FreezeResetBegin = True
            If Debug_mouse Then
                GdipSetSolidFillColor ECore.pB, argb(255, 27, 27, 27)
                GdipFillEllipse ECore.UPage.GG, ECore.pB, X - 10, y - 10, 20, 20
                GdipSetSolidFillColor ECore.pB, argb(80, 255, 0, 0)
                GdipFillRectangle ECore.UPage.GG, ECore.pB, X, y, w, h
                EF.Writes Mouse.State + 1, X - 10, y - 7, ECore.UPage.GG, argb(255, 255, 255, 255), 14, 20, 0, StringAlignmentCenter, FontStyleBold
            End If
            CheckMouse = Mouse.State + 1
            If Mouse.State = 2 Then Mouse.State = 0
        End If
    End Function
    Public Function CheckMouse2() As MButtonState
        'Return Value:0=none,1=in,2=down,3=up
        If Debug_mouse Then
            GdipSetSolidFillColor ECore.pB, argb(20, 0, 0, 255)
            GdipFillRectangle ECore.UPage.GG, ECore.pB, DrawF.X, DrawF.y, DrawF.Width, DrawF.Height
        End If
        If ECore.LockPage <> "" Then
            If ECore.LockPage <> ECore.UpdatingPage Then Exit Function
        End If
        If Mouse.X >= DrawF.X And Mouse.y >= DrawF.y And Mouse.X <= DrawF.X + DrawF.Width And Mouse.y <= DrawF.y + DrawF.Height Then
            If ECore.FreezeMode Then ECore.FreezeResetBegin = True
            If Debug_mouse Then
                GdipSetSolidFillColor ECore.pB, argb(255, 27, 27, 27)
                GdipFillEllipse ECore.UPage.GG, ECore.pB, DrawF.X - 10, DrawF.y - 10, 20, 20
                GdipSetSolidFillColor ECore.pB, argb(80, 0, 0, 255)
                GdipFillRectangle ECore.UPage.GG, ECore.pB, DrawF.X, DrawF.y, DrawF.Width, DrawF.Height
                EF.Writes Mouse.State + 1, DrawF.X - 10, DrawF.y - 7, ECore.UPage.GG, argb(255, 255, 255, 255), 14, 20, 0, StringAlignmentCenter, FontStyleBold
            End If
            CheckMouse2 = Mouse.State + 1
            If DrawF.CrashIndex <> 0 Then
                If ColorLists(DrawF.CrashIndex).IsAlpha((Mouse.X - DrawF.X) * DrawF.WSc, (Mouse.y - DrawF.y) * DrawF.HSc) = False Then CheckMouse2 = mMouseOut: Exit Function
            End If
            If Mouse.State = 2 Then Mouse.State = 0
        End If
    End Function
'========================================================
'   KeyBoard
    Public Function IsKeyPress(Code As Long) As Boolean
        IsKeyPress = (GetAsyncKeyState(Code) < 0)
    End Function
    Public Function IsKeyUp(Code As Long) As Boolean
        Dim t As Boolean
        t = LastKeyUpRet
        LastKeyUpRet = (GetAsyncKeyState(Code) < 0)
        If t = True And LastKeyUpRet = False Then IsKeyUp = True
    End Function
'========================================================
'   Screen Window
    Public Function StartScreenDialog(w As Long, h As Long, ch As Object) As Object
        Set StartScreenDialog = New EmeraldWindow
        StartScreenDialog.NewFocusWindow w, h, ch
        Dim f As Object
        For Each f In VB.Forms
            If TypeName(f) <> "EmeraldWindow" Then f.Enabled = False
        Next
    End Function
'========================================================
'   AssetsTree
    Public Function AddAssetsTree(Tree As AssetsTree, arg1 As Variant, arg2 As Variant)
        ReDim Preserve AssetsTrees(UBound(AssetsTrees) + 1)
        AssetsTrees(UBound(AssetsTrees)) = Tree
        AssetsTrees(UBound(AssetsTrees)).arg1 = arg1
        AssetsTrees(UBound(AssetsTrees)).arg2 = arg2
    End Function
    Public Function FindAssetsTree(path As String, arg1 As Variant, arg2 As Variant) As Integer
        On Error Resume Next
        For I = 1 To UBound(AssetsTrees)
            If AssetsTrees(I).path = path And AssetsTrees(I).arg1 = arg1 And AssetsTrees(I).arg2 = arg2 Then
                If Err.Number <> 0 Then
                    Err.Clear
                Else
                    FindAssetsTree = I: Exit For
                End If
            End If
        Next
    End Function
    Public Function GetAssetsTree(path As String) As AssetsTree
        For I = 1 To UBound(AssetsTrees)
            If AssetsTrees(I).path = path Then GetAssetsTree = AssetsTrees(I): Exit For
        Next
    End Function
'========================================================

Attribute VB_Name = "GCore"
'========================================================
'   Emerald 绘图框架模块
'   更新内容(ver.501)
'   -添加IsKeyUp
'   -新增开场LOGO设置
'   -添加版本更新注意事项
'   -添加过场效果枚举
'   -修复部分字体样式不可用的问题
'   -添加版本更新检测
'   更新内容(ver.430)
'   -添加ImgCount，ImgSize
'   -添加启动页面
'   -添加资源加载进度条
'   -修复可以以非法字符开头创建工程的错误
'   -添加模板代码
'   -现在卸载过程中途发生错误会继续执行接下来的步骤
'   -修复按下关闭按钮默认选择确认的问题
'   更新内容(ver.427)
'   -修复Builder零错误问题
'   -修改图标使其能够被清晰地察觉
'   -现在在更新工程前需要更新Builder
'   -修复不能卸载的问题（Builder）
'   -注册软件信息（Builder）
'   -添加圆角矩形
'   -添加DrawImageEx
'   -新增Builder GUI
'   更新内容(ver.426)
'   -修复矩形边框柔和边缘的问题
'   -添加画笔大小
'   -现在授权界面可以使用鼠标
'   -添加鼠标枚举
'   -全新Debug界面
'   更新内容(ver.420)
'   -增加屏幕窗口
'   -修复矩形柔和边缘的问题
'   更新内容(ver.419)
'   -添加过程一键建立器
'   -修复滑轮方向错误的问题
'   -添加文件名检测（Builder）
'   -添加最新文件更新（Builder）
'   -现在可以卸载Builder
'   -现在Builder的提示文本是中文
'   -现在Builder默认需要管理员权限
'   更新内容(ver.329)
'   -添加碰撞箱
'   -添加调试功能
'   更新内容(ver.324)
'   -完善动画
'   -修复paint动画异常的问题
'   -添加动画播放完成的事件
'   -添加动画延迟属性
'   更新内容(ver.323)
'   -DPI适应
'   -鼠标滚轮支持
'   更新内容(ver.317)
'   -新增窗口失焦和取得焦点事件
'   更新内容(ver.316)
'   -修复卷轴模式的一些问题
'   -新增6种过场特效
'   -新增动画
'   -新增页面切换事件
'   更新内容(ver.315)
'   -新增卷轴模式
'   -现在可以检测键盘按下
'   -现在支持动态GIF图片
'   -新增四种过场特效
'   更新日志(ver.211)
'   -添加窗口模糊方法（Blurto）
'========================================================
'   DPI适应
    Public Declare Function SetProcessDpiAwareness Lib "SHCORE.DLL" (ByVal DPImodel As Long) As Long
'=========================================================================
    Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
    Public Type MState
        State As Integer
        button As Integer
        X As Single
        Y As Single
    End Type
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
    End Enum
    Public ECore As GMan, EF As GFont, EAni As Object, ESave As GSaving
    Public GHwnd As Long, GDC As Long, GW As Long, GH As Long
    Public Mouse As MState, DrawF As RECT
    Public FPS As Long, FPSt As Long, tFPS As Long, FPSct As Long, FPSctt As Long
    Public SysPage As GSysPage
    Public PreLoadCount As Long, LoadedCount As Long
    Public Const Version As Long = 19050106
    Dim LastKeyUpRet As Boolean
    Dim Wndproc As Long
'========================================================
'   Init
    Public Sub StartEmerald(Hwnd As Long, W As Long, H As Long)
    
        If DebugMode Then
            If App.LogMode <> 0 Then MsgBox "错误：生成时未关闭Debug模式。": End
        End If
        
        InitGDIPlus
        BASS_Init -1, 44100, BASS_DEVICE_3D, Hwnd, 0
        GHwnd = Hwnd: GW = W: GH = H
        SetWindowPos Hwnd, 0, 0, 0, W, H + 30, SWP_NOMOVE Or SWP_NOZORDER
        GDC = GetDC(Hwnd)
        If App.LogMode <> 0 Then Wndproc = SetWindowLongA(Hwnd, GWL_WNDPROC, AddressOf Process)
        
        Set EAni = New GAnimation
        Set SysPage = New GSysPage
        
        If Val(GetWinNTVersion) > 6.1 Then               '如果当前系统版本高于win7
            SetProcessDpiAwareness 2&                    '调用API使本程序在高DPI情况下不模糊
        End If
        
        If DebugMode Then
            DebugWindow.Show
        End If
        
        If App.LogMode = 0 Then Call CheckUpdate
    End Sub
    Public Sub EndEmerald()
        If DebugMode Then
            Unload DebugWindow
        End If
        
        If App.LogMode <> 0 Then SetWindowLongA GHwnd, GWL_WNDPROC, Wndproc
        If Not (ECore Is Nothing) Then ECore.Dispose
        If Not (EF Is Nothing) Then EF.Dispose
        TerminateGDIPlus
        BASS_Free
    End Sub
    Public Sub MakeFont(ByVal name As String)
        Set EF = New GFont
        EF.MakeFont name
    End Sub
'========================================================
'   RunTime
    Public Function Process(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        On Error GoTo sth

        If uMsg = WM_MOUSEWHEEL Then
            Dim Direction As Integer, Strong As Single
            Direction = IIf(wParam < 0, -1, 1): Strong = Abs(wParam / 7864320)
            ECore.Wheel Direction, Strong
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
        Dim i As Long, g As Long, e As Long, b As BlurParams, W As Long, H As Long
        '粘贴到缓冲窗口
        buffWin.AutoRedraw = True
        BitBlt buffWin.hdc, 0, 0, GW, GH, srcDC, 0, 0, vbSrcCopy: buffWin.Refresh
        
        '创建Bitmap
        GdipCreateBitmapFromHBITMAP buffWin.Image.handle, buffWin.Image.hpal, i
        
        '模糊操作
        GdipCreateEffect2 GdipEffectType.Blur, e: b.Radius = Radius: GdipSetEffectParameters e, b, LenB(b)
        GdipGetImageWidth i, W: GdipGetImageHeight i, H
        GdipBitmapApplyEffect i, e, NewRectL(0, 0, W, H), 0, 0, 0
        
        '画~
        GdipCreateFromHDC DC, g
        GdipDrawImage g, i, 0, 0
        GdipDisposeImage i: GdipDeleteGraphics g: GdipDeleteEffect e '垃圾处理
        buffWin.AutoRedraw = False
    End Sub
    Public Sub BlurImg(img As Long, Radius As Long)
        Dim b As BlurParams, e As Long, W As Long, H As Long
        
        '模糊操作
        GdipCreateEffect2 GdipEffectType.Blur, e: b.Radius = Radius: GdipSetEffectParameters e, b, LenB(b)
        GdipGetImageWidth img, W: GdipGetImageHeight img, H
        GdipBitmapApplyEffect img, e, NewRectL(0, 0, W, H), 0, 0, 0
        
        '画~
        GdipDeleteEffect e '垃圾处理
    End Sub
    Public Function CreateCDC(W As Long, H As Long) As Long
        Dim bm As BITMAPINFOHEADER, DC As Long, DIB As Long
    
        With bm
            .biBitCount = 32
            .biHeight = H
            .biWidth = W
            .biPlanes = 1
            .biSizeImage = (.biWidth * .biBitCount + 31) / 32 * 4 * .biHeight
            .biSize = Len(bm)
        End With
        
        DC = CreateCompatibleDC(GDC)
        DIB = CreateDIBSection(DC, bm, DIB_RGB_COLORS, ByVal 0, 0, 0)
        DeleteObject SelectObject(DC, DIB)
        
        CreateCDC = DC
    End Function
    Public Sub PaintDC(DC As Long, destDC As Long, Optional X As Long = 0, Optional Y As Long = 0, Optional cx As Long = 0, Optional cy As Long = 0, Optional cw, Optional ch, Optional Alpha)
        Dim b As BLENDFUNCTION, Index As Integer, bl As Long
        
        If Not IsMissing(Alpha) Then
            If Alpha < 0 Then Alpha = 0
            If Alpha > 1 Then Alpha = 1
            With b
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(Alpha * 255)
            End With
            CopyMemory bl, b, 4
        End If
        
        If IsMissing(cw) Then cw = GW - cx
        If IsMissing(ch) Then ch = GH - cy
        
        If IsMissing(Alpha) Then
            BitBlt destDC, X, Y, cw, ch, DC, cx, cy, vbSrcCopy
        Else
            AlphaBlend destDC, X, Y, cw, ch, DC, cx, cy, cw, ch, bl
        End If
    End Sub
    Function Cubic(t As Single, arg0 As Single, arg1 As Single, arg2 As Single, arg3 As Single) As Single
        'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
        'Attention:all the args must in this area (0~1)
        Cubic = (arg0 * ((1 - t) ^ 3)) + (3 * arg1 * t * ((1 - t) ^ 2)) + (3 * arg2 * (t ^ 2) * (1 - t)) + (arg3 * (t ^ 3))
    End Function
'========================================================
'   Mouse
    Public Sub UpdateMouse(X As Single, Y As Single, State As Long, button As Integer)
        With Mouse
            .X = X
            .Y = Y
            .State = State
            .button = button
        End With
    End Sub
    Public Function CheckMouse(X As Long, Y As Long, W As Long, H As Long) As MButtonState
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.X >= X And Mouse.Y >= Y And Mouse.X <= X + W And Mouse.Y <= Y + H Then
            CheckMouse = Mouse.State + 1
            If Mouse.State = 2 Then Mouse.State = 0
        End If
    End Function
    Public Function CheckMouse2() As MButtonState
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.X >= DrawF.Left And Mouse.Y >= DrawF.top And Mouse.X <= DrawF.Left + DrawF.Right And Mouse.Y <= DrawF.top + DrawF.Bottom Then
            CheckMouse2 = Mouse.State + 1
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
    Public Function StartScreenDialog(W As Long, H As Long, ch As Object) As Object
        Set StartScreenDialog = New EmeraldWindow
        StartScreenDialog.NewFocusWindow W, H, ch
        Dim f As Object
        For Each f In VB.Forms
            If TypeName(f) <> "EmeraldWindow" Then f.Enabled = False
        Next
    End Function
'========================================================
'   Update
    Public Sub CheckUpdate()
        On Error Resume Next
        Dim Data As New GSaving
        Data.Create "Emerald.Core", "Emerald.Core"
        If Now - CDate(Data.GetData("UpdateTime")) >= UpdateCheckInterval Or Data.GetData("UpdateAble") = 1 Then
            Data.PutData "UpdateTime", Now
            
            Dim XmlHttp As Object, ret As String
            Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
            XmlHttp.Open "GET", "https://raw.githubusercontent.com/Red-Error404/Emerald/master/Version.txt", True
            XmlHttp.Send
            Do While XmlHttp.ReadyState <> 4
                DoEvents
            Loop
            ret = XmlHttp.responseText
            Set XmlHttp = Nothing
            Debug.Print Now, "检查版本完毕，最新版本号：" & ret
            
            If Val(ret) <> Version Then
                Data.PutData "UpdateAble", 1
                If MsgBox("发现Emerald存在新版本，您希望现在前往下载吗？", vbYesNo + 48, "Emerald") = vbNo Then Exit Sub
                Data.PutData "UpdateAble", 0
                ShellExecuteA 0, "open", "https://github.com/Red-Error404/Emerald", "", "", SW_SHOW
            End If
        End If
    End Sub
'========================================================

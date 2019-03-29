Attribute VB_Name = "GCore"
'========================================================
'   Emerald 绘图框架模块
'   更新内容(ver.329)
'   -添加碰撞箱
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
    x As Single
    y As Single
End Type
Public ECore As GMan, EF As GFont, EAni As Object
Public GHwnd As Long, GDC As Long, GW As Long, GH As Long
Public Mouse As MState, DrawF As RECT
Dim Wndproc As Long
'========================================================
'   Init
    Public Sub StartEmerald(hwnd As Long, w As Long, h As Long)
        InitGDIPlus
        BASS_Init -1, 44100, BASS_DEVICE_3D, hwnd, 0
        GHwnd = hwnd: GW = w: GH = h
        GDC = GetDC(hwnd)
        If App.LogMode <> 0 Then Wndproc = SetWindowLongA(hwnd, GWL_WNDPROC, AddressOf Process)
        
        Set EAni = New GAnimation
        
        If Val(GetWinNTVersion) > 6.1 Then               '如果当前系统版本高于win7
            SetProcessDpiAwareness 2&                    '调用API使本程序在高DPI情况下不模糊
        End If
    End Sub
    Public Sub EndEmerald()
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
    Public Function Process(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        On Error GoTo sth

        If uMsg = WM_MOUSEWHEEL Then
            Dim Direction As Integer, Strong As Single
            Direction = IIf(wParam < 0, 1, -1): Strong = Abs(wParam / 7864320)
            ECore.Wheel Direction, Strong
        End If
        
last:
        Process = CallWindowProcA(Wndproc, hwnd, uMsg, wParam, lParam)
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
    Public Sub BlurTo(dc As Long, srcDC As Long, buffWin As Form, Optional Radius As Long = 60)
        Dim i As Long, g As Long, e As Long, b As BlurParams, w As Long, h As Long
        '粘贴到缓冲窗口
        buffWin.AutoRedraw = True
        BitBlt buffWin.hdc, 0, 0, GW, GH, srcDC, 0, 0, vbSrcCopy: buffWin.Refresh
        
        '创建Bitmap
        GdipCreateBitmapFromHBITMAP buffWin.Image.handle, buffWin.Image.hpal, i
        
        '模糊操作
        GdipCreateEffect2 GdipEffectType.Blur, e: b.Radius = Radius: GdipSetEffectParameters e, b, LenB(b)
        GdipGetImageWidth i, w: GdipGetImageHeight i, h
        GdipBitmapApplyEffect i, e, NewRectL(0, 0, w, h), 0, 0, 0
        
        '画~
        GdipCreateFromHDC dc, g
        GdipDrawImage g, i, 0, 0
        GdipDisposeImage i: GdipDeleteGraphics g: GdipDeleteEffect e '垃圾处理
        buffWin.AutoRedraw = False
    End Sub
    Public Function CreateCDC(w As Long, h As Long) As Long
        Dim bm As BITMAPINFOHEADER, dc As Long, DIB As Long
    
        With bm
            .biBitCount = 32
            .biHeight = h
            .biWidth = w
            .biPlanes = 1
            .biSizeImage = (.biWidth * .biBitCount + 31) / 32 * 4 * .biHeight
            .biSize = Len(bm)
        End With
        
        dc = CreateCompatibleDC(GDC)
        DIB = CreateDIBSection(dc, bm, DIB_RGB_COLORS, ByVal 0, 0, 0)
        DeleteObject SelectObject(dc, DIB)
        
        CreateCDC = dc
    End Function
    Public Sub PaintDC(dc As Long, destDC As Long, Optional x As Long = 0, Optional y As Long = 0, Optional cx As Long = 0, Optional cy As Long = 0, Optional cw, Optional ch, Optional alpha)
        Dim b As BLENDFUNCTION, index As Integer, bl As Long
        
        If Not IsMissing(alpha) Then
            If alpha < 0 Then alpha = 0
            If alpha > 1 Then alpha = 1
            With b
                .AlphaFormat = &H1
                .BlendFlags = &H0
                .BlendOp = 0
                .SourceConstantAlpha = Int(alpha * 255)
            End With
            CopyMemory bl, b, 4
        End If
        
        If IsMissing(cw) Then cw = GW - cx
        If IsMissing(ch) Then ch = GH - cy
        
        If IsMissing(alpha) Then
            BitBlt destDC, x, y, cw, ch, dc, cx, cy, vbSrcCopy
        Else
            AlphaBlend destDC, x, y, cw, ch, dc, cx, cy, cw, ch, bl
        End If
    End Sub
    Function Cubic(t As Single, arg0 As Single, arg1 As Single, arg2 As Single, arg3 As Single) As Single
        'Formula:B(t)=P_0(1-t)^3+3P_1t(1-t)^2+3P_2t^2(1-t)+P_3t^3
        'Attention:all the args must in this area (0~1)
        Cubic = (arg0 * ((1 - t) ^ 3)) + (3 * arg1 * t * ((1 - t) ^ 2)) + (3 * arg2 * (t ^ 2) * (1 - t)) + (arg3 * (t ^ 3))
    End Function
'========================================================
'   Mouse
    Public Sub UpdateMouse(x As Single, y As Single, State As Long, button As Integer)
        With Mouse
            .x = x
            .y = y
            .State = State
            .button = button
        End With
    End Sub
    Public Function CheckMouse(x As Long, y As Long, w As Long, h As Long) As Integer
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.x >= x And Mouse.y >= y And Mouse.x <= x + w And Mouse.y <= y + h Then
            CheckMouse = Mouse.State + 1
            If Mouse.State = 2 Then Mouse.State = 0
        End If
    End Function
    Public Function CheckMouse2() As Integer
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.x >= DrawF.Left And Mouse.y >= DrawF.top And Mouse.x <= DrawF.Left + DrawF.Right And Mouse.y <= DrawF.top + DrawF.Bottom Then
            CheckMouse2 = Mouse.State + 1
            If Mouse.State = 2 Then Mouse.State = 0
        End If
    End Function
'========================================================
'   KeyBoard
    Public Function IsKeyPress(Code As Long) As Boolean
        IsKeyPress = (GetAsyncKeyState(Code) < 0)
    End Function
'========================================================

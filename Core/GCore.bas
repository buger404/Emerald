Attribute VB_Name = "GCore"
'========================================================
'   Emerald 绘图框架模块
'   更新内容(ver.315)
'   -新增卷轴模式
'   -现在可以检测键盘按下
'   -现在支持动态GIF图片
'   -新增四种过场特效
'   更新日志(ver.211)
'   -添加窗口模糊方法（Blurto）
'========================================================
Private Declare Sub AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal hHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal BLENDFUNCTION As Long) ' As Long
Public Type MState
    State As Integer
    button As Integer
    X As Single
    Y As Single
End Type
Public ECore As GMan, EF As GFont
Public GHwnd As Long, GDC As Long, GW As Long, GH As Long
Public Mouse As MState, DrawF As RECT
'========================================================
'   Init
    Public Sub StartEmerald(hwnd As Long, W As Long, H As Long)
        InitGDIPlus
        BASS_Init -1, 44100, BASS_DEVICE_3D, hwnd, 0
        GHwnd = hwnd: GW = W: GH = H
        GDC = GetDC(hwnd)
    End Sub
    Public Sub EndEmerald()
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
    Public Sub BlurTo(dc As Long, srcDC As Long, buffWin As Form, Optional Radius As Long = 60)
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
        GdipCreateFromHDC dc, g
        GdipDrawImage g, i, 0, 0
        GdipDisposeImage i: GdipDeleteGraphics g: GdipDeleteEffect e '垃圾处理
        buffWin.AutoRedraw = False
    End Sub
    Public Function CreateCDC(W As Long, H As Long) As Long
        Dim bm As BITMAPINFOHEADER, dc As Long, DIB As Long
    
        With bm
            .biBitCount = 32
            .biHeight = H
            .biWidth = W
            .biPlanes = 1
            .biSizeImage = (.biWidth * .biBitCount + 31) / 32 * 4 * .biHeight
            .biSize = Len(bm)
        End With
        
        dc = CreateCompatibleDC(GDC)
        DIB = CreateDIBSection(dc, bm, DIB_RGB_COLORS, ByVal 0, 0, 0)
        DeleteObject SelectObject(dc, DIB)
        
        CreateCDC = dc
    End Function
    Public Sub PaintDC(dc As Long, destDC As Long, Optional X As Long = 0, Optional Y As Long = 0, Optional cx As Long = 0, Optional cy As Long = 0, Optional cw, Optional ch, Optional alpha)
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
            BitBlt destDC, X, Y, cw, ch, dc, cx, cy, vbSrcCopy
        Else
            AlphaBlend destDC, X, Y, cw, ch, dc, cx, cy, cw, ch, bl
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
    Public Function CheckMouse(X As Long, Y As Long, W As Long, H As Long) As Integer
        'Return Value:0=none,1=in,2=down,3=up
        If Mouse.X >= X And Mouse.Y >= Y And Mouse.X <= X + W And Mouse.Y <= Y + H Then
            CheckMouse = Mouse.State + 1
            If Mouse.State = 2 Then Mouse.State = 0
        End If
    End Function
    Public Function CheckMouse2() As Integer
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
'========================================================

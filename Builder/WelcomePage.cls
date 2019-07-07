VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WelcomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
'======================================
    Dim HideTime As Long
    'Dim test As String
Public Sub DrawSetup()

    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布

    Page.Paint 0, 0, 150, GW, GH - 150, argb(255, 255, 255, 255)
    
    If Page.Res.GetImage("setupappicon.png") = 0 Then
        Page.Paint 1, 50, 150 - 64, 128, 128, argb(255, 95, 229, 221)
        Page.Writes "G", 50 + 64, 150 - 64 + 32, 48, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Else
        Page.DrawImage "setupappicon.png", 50, 150 - 64, alpha:=1
    End If
    
    Page.Writes SPackage.GameName, GW - 50, 97, 24, argb(255, 27, 27, 27), style:=FontStyleBold, align:=StringAlignmentFar
    Page.Writes SPackage.GameVersion, GW - 50, 170, 18, argb(150, 27, 27, 27), align:=StringAlignmentFar
    Page.Writes "By " & SPackage.Maker, GW - 50, 200, 16, argb(255, 95, 229, 221), align:=StringAlignmentFar
    
    Page.Writes "简介：" & vbCrLf & SPackage.GameDescribe, 50, 250, 18, argb(150, 27, 27, 27)
    
    Page.Writes "安装路径", 49, GH - 110, 16, argb(150, 27, 27, 27)
    If Page.ShowEdit(SSetupPath, 0, 130, GH - 110, 450, 25, argb(120, 27, 27, 27), argb(0, 0, 0, 0), argb(255, 242, 242, 242), argb(255, 85, 219, 211), size:=16) = 4 Then
        If Right(SSetupPath, 1) = "\" Then SSetupPath = Left(SSetupPath, Len(SSetupPath) - 1)
        Dim DadPath As String
        DadPath = Left(SSetupPath, Len(SSetupPath) - Len(Split(SSetupPath, "\")(UBound(Split(SSetupPath, "\")))))
        If Dir(DadPath, vbDirectory) = "" Then
            SSetupPath = "C:\Program Files\" & SPackage.GameName
            Dialog "错误", "指定的安装位置无效，因为父级目录不存在。", "哦"
        End If
    End If
    
    Page.ShowColorCheckBox LnkSwitch, 50, GH - 70, 100, 23, "创建桌面快捷方式", argb(255, 158, 158, 158), argb(255, 85, 219, 211), argb(150, 27, 27, 27), 16, FontStyleRegular
    
    If Page.ShowColorButton(0, GW - 170, GH - 80, 130, 40, "安装", argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
        ECore.NewTransform transFadeIn, 700, "SetupPage"
        CmdMark = "Setup"
        Call SetupPack
    End If
    
End Sub
Public Sub DrawSet()

    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布

    Page.Writes "设置", GW / 2, 100, Color:=argb(255, 27, 27, 27), size:=32, align:=StringAlignmentCenter, animation:=1
    Page.Writes "在这里管理全局设置", GW / 2, 150, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
    
    Page.ShowColorCheckBox DebugMode, 50, 210, 300, 24, "开启Debug调试栏", argb(255, 222, 222, 222), argb(255, 95, 229, 221), argb(150, 27, 27, 27), size:=16
    Page.ShowColorCheckBox HideLOGO, 50, 250, 300, 24, "禁用LOGO动画", argb(255, 222, 222, 222), argb(255, 95, 229, 221), argb(150, 27, 27, 27), size:=16
    Page.ShowColorCheckBox DisableLOGO, 50, 290, 300, 24, "省略多余的LOGO动画", argb(255, 222, 222, 222), argb(255, 95, 229, 221), argb(150, 27, 27, 27), size:=16
    Page.Writes "更新间隔： " & UpdateCheckInterval & "（天）", 50, 350, 16, argb(150, 27, 27, 27), 300, 24
    If CheckMouse2 = mMouseUp Then UpdateCheckInterval = Val(InputBox("输入新的值：" & vbCrLf & "UpdateCheckInterval", "Emerald", UpdateCheckInterval))
    Page.Writes "更新超时： " & UpdateTimeOut & "（毫秒）", 50, 380, 16, argb(150, 27, 27, 27), 300, 24
    If CheckMouse2 = mMouseUp Then UpdateTimeOut = Val(InputBox("输入新的值：" & vbCrLf & "UpdateTimeOut", "Emerald", UpdateTimeOut))

    If Page.ShowColorButton(0, GW / 2 - 130 / 2, GH - 80, 130, 40, "保存", argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
        Dim data As New GSaving
        data.Create "Emerald.Core"
        data.AutoSave = True
        Call SaveSettings(data)
        Set data = Nothing
        
        SetMode = False
        ECore.NewTransform , 700, "WelcomePage"
    End If
End Sub
Public Sub Update()
    '绘制事件
    
    If SetMode Then Call DrawSet: Exit Sub
    If SetupMode Then Call DrawSetup: Exit Sub
    
    Dim BtnStr As String
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布

    If Repaired Then
        Page.Writes "Something Wrong", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
        Page.Writes "There is something wrong with your Builder .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
        BtnStr = "Repair"
        GoTo Draw
    End If
    
    If InstalledPath <> "" Then
        If (Not IsUpdate) Then
            If Cmd = "-uninstall" Then
                Page.Writes "Say Goodbye to Emerald Builder", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
                Page.Writes "Are you really going to uninstall me ?", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
                BtnStr = "Uninstall"
            Else
            
                Abouting = True
            
                Dim pro As Single
                If HideTime <> 0 Then pro = Cubic((GetTickCount - 3000 - HideTime) / 1000, 0, 1, 1, 1)
                If pro > 1 Then pro = 1
                If pro < 0 Then pro = 0
            
                Page.Clear argb(255 * 0.85, 64, 64, 73) '清空画布
                
                'Page.DrawAnimation "test", 250, 250
                'If CheckMouse2 = mMouseUp Then
                '    Page.SwitchChannel "test", IIf(Mouse.button = 1, "Default", "Test2")
                'End If
                'Page.DrawAnimation "test", 300, 250
                'Page.DrawAnimation "test", 200, 250
                'Page.DrawAnimation "test", 400, 250
                'Page.DrawAnimation "test", 100, 0
                
                'If Page.ShowEdit(test, 2, 100, 100, 300, 40, argb(255, 0, 0, 0), argb(180, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 210, 210, 210), 64, 0, 28) = 4 Then
                '    MsgBox "文本框内容：" & test
                'End If
                'If Page.ShowEdit(test, 0, 100, 160, 300, 40, argb(255, 0, 0, 0), argb(180, 255, 255, 255), argb(255, 255, 255, 255), argb(255, 210, 210, 210), 0, 0, 28) = 4 Then
                '    MsgBox "文本框内容：" & test
                'End If
                
                Page.DrawImage "404.png", 165, GH / 2, alpha:=1 - pro, pos:=posOnCenter, animation:=1

                'Page.ShowColorCheckBox Debug_focus, 20, 100, 200, 24, "Debug_Switch", argb(255, 222, 222, 222), argb(255, 0, 176, 240), argb(150, 242, 242, 242), 18
                
                Page.Writes "关于", 350 + pro * 22, 170 + pro * 36, Color:=argb(255, 242, 242, 242), size:=18 + pro * 4, animation:=1, style:=FontStyleBold
                Page.Writes "制作： Error 404（QQ 1361778219）" & vbCrLf & "邮箱： ris_vb@126.com", 350 - pro * 110, 220 + pro * 40, Color:=argb(150, 242, 242, 242), size:=18, animation:=2
                Page.Writes "版本： " & Version, 350 - pro * 50, 270 + pro * 40, Color:=argb(150 - pro * 150, 242, 242, 242), size:=18, animation:=2
                Page.Writes "感谢Inter.net群，Redstone群和南笙的支持。", 350 - pro * 110, 340 + pro * 180, Color:=argb(150 - pro * 60, 242, 242, 242), size:=18 - pro * 2, animation:=2
                Page.Writes "Emerald 开发文档", 350 - pro * 30, 370 - pro * 40, w:=200, h:=25, Color:=argb(255, 255, 255, 255), size:=18, animation:=2, style:=FontStyleBold
                If CheckMouse2 = mMouseUp Then
                    ShellExecuteA 0, "open", "https://github.com/buger404/Emerald/wiki", "", "", SW_SHOW
                End If
                
                If NewVersion = 0 Then
                    NewVersion = 1
                    Call CheckOnLineUpdate
                End If
                
                Select Case NewVersion
                    Case 1
                        Page.ShowLoading 30, GH - 30 - 24, 24, 24, 3, argb(255, 242, 195, 17), argb(255, 240, 127, 34), argb(255, 232, 76, 61)
                        Page.Writes "正在检查更新...", 80, GH - 30 - 24, Color:=argb(150, 242, 242, 242), size:=18, animation:=2
                    Case Is > Version
                        If HideTime = 0 Then HideTime = GetTickCount
                        Page.Paint 1, 30, GH - 30 - 24, 24, 24, argb(255, 255, 206, 69)
                        Page.Writes "有更新可用（" & NewVersion & "），点击此处下载。", 80, GH - 30 - 24, w:=GW, h:=25, Color:=argb(255, 232, 76, 61), size:=18, animation:=2
                        If CheckMouse2 = mMouseUp Then ShellExecuteA 0, "open", "https://github.com/buger404/Emerald", "", "", SW_SHOW
                    Case Is <= Version
                        If HideTime = 0 Then HideTime = GetTickCount
                        Page.Paint 1, 30, GH - 30 - 24, 24, 24, argb(255 - pro * 255, 95, 229, 221)
                        Page.Writes "已经是最新版本", 80, GH - 30 - 24, Color:=argb(150 - pro * 150, 242, 242, 242), size:=18, animation:=2
                    Case 3
                        If HideTime = 0 Then HideTime = GetTickCount
                        Page.Paint 1, 30, GH - 30 - 24, 24, 24, argb(255 - pro * 255, 254, 84, 57)
                        Page.Writes "连接超时", 80, GH - 30 - 24, Color:=argb(150 - pro * 150, 242, 242, 242), size:=18, animation:=2
                End Select
                
                Exit Sub
            End If
        Else
            Page.Writes "New Version : " & Version, GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
            Page.Writes "Make everything better .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
            BtnStr = "Update"
        End If
    End If
    
    If InstalledPath = "" Then
        Page.Writes "Welcome to Emerald Builder", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
        Page.Writes "Let's create our own games together .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
        BtnStr = "Setup"
    End If
        
Draw:
    If Page.ShowColorButton(0, GW / 2 - 130 / 2, GH - 150, 130, 40, BtnStr, argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
        If BtnStr = "Uninstall" Then
            If Dialog("卸载", "你确定要卸载Emerald Builder？", "是的", "手滑") <> 1 Then Exit Sub
            ECore.NewTransform transFadeIn, 700, "SetupPage"
            CmdMark = BtnStr
            Call Uninstall
        Else
            ECore.NewTransform transFadeIn, 700, "SetupPage"
            CmdMark = BtnStr
            Call Setup
        End If
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub Enter()
    '页面进入事件
    Page.StartAnimation 1
    Page.StartAnimation 2, 200
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    If PackPos = -1 Then
        Page.Res.NewImages App.Path & "\assets"
        Page.Res.NewImages App.Path & "\assets\debug"
        '导入动画
        'Page.LoadAnimationsFromDir App.path & "\animation"
        'Page.CreatePlayAnimation   "test", "test", "Default"
    End If
    Dim m(4, 4) As Single
    m(0, 0) = 1: m(0, 1) = 0: m(0, 2) = 0: m(0, 3) = 0
    m(1, 0) = 0: m(1, 1) = 0: m(1, 2) = 0: m(1, 3) = 0
    m(2, 0) = 0: m(2, 1) = 0: m(2, 2) = 0: m(2, 3) = 0
    m(3, 0) = 1: m(3, 1) = 0: m(3, 2) = 0: m(3, 3) = 1
    
    'Page.Res.ClipCircle "404.png"
    '笔记
    'Page.Res.ApplyBlurEffect 模糊半径，四周阴影
    'ApplyBrightnessContrastEffect 亮度，对比度
    'ApplyColorBalanceEffect 红色调，绿色调，蓝色调
    'ApplyColorCurveEffect 对指定颜色通道应用某个特殊效果
    'ApplyColorMatrixEffect 反正颜色矩阵就对了
    'ApplyHueSaturationLightnessEffect 色相，亮度，饱和度
    'ApplyLevelsEffect 高光，中间色，暗调
    'ApplySharpenEffect 数量，锐化半径
    'ApplyTintEffect 程度，色调
    
    '创建页面
    ECore.Add Page, "WelcomePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
        Page.NewAnimation False, "FadeIn", 2, 500
        Page.NewAnimation False, "FadeIn", 2, 500
    '===============================================
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub

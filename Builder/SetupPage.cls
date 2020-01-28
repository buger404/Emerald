VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SetupPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
    Public Step As Long, SetupInfo As String
    Public Progress As Single
    Public ShowLog As Boolean, LaunchNow As Boolean
'======================================
Public Sub DrawStep5()
    Page.ShowColorCheckBox ShowLog, GW / 2 - 100, GH / 2 - 20, 200, 20, "显示卸载程序报告", argb(255, 198, 198, 198), argb(255, 35, 170, 242), argb(255, 64, 64, 64), 18

    If Page.ShowColorButton(0, GW / 2 - 70, GH - 120, 140, 40, "完成", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=18) = 3 Then
        If ShowLog Then ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
        Unload MainWindow
        End
    End If
    'ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
End Sub
Public Sub DrawStep4()
    Page.ShowLoading GW / 2 - 100, GH / 2 - 160, 200, 200, 9, argb(255, 35, 170, 242), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
    Page.Writes SetupInfo, GW / 2, GH - 130, 16, argb(255, 32, 32, 32), align:=StringAlignmentCenter
    Page.Writes format(Int(Progress * 1000) / 10, "0.0") & "%", GW / 2, GH - 100, 16, argb(255, 32, 32, 32), align:=StringAlignmentCenter, style:=FontStyleBold
End Sub
Public Sub DrawStep3()
    Page.ShowColorCheckBox LaunchNow, GW / 2 - 100, GH / 2 - 60, 200, 20, "立即启动", argb(255, 198, 198, 198), argb(255, 35, 170, 242), argb(255, 64, 64, 64), 18
    Page.ShowColorCheckBox ShowLog, GW / 2 - 100, GH / 2 - 20, 200, 20, "显示安装程序报告", argb(255, 198, 198, 198), argb(255, 35, 170, 242), argb(255, 64, 64, 64), 18

    If Page.ShowColorButton(0, GW / 2 - 70, GH - 120, 140, 40, "完成安装", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=18) = 3 Then
        If LaunchNow Then ShellExecuteA 0, "open", SSetupPath & "\app.exe", "", "", SW_SHOW
        If ShowLog Then ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
        Unload MainWindow
        End
    End If
    'ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
End Sub
Public Sub DrawStep2()
    Page.ShowLoading GW / 2 - 100, GH / 2 - 160, 200, 200, 9, argb(255, 35, 170, 242), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
    Page.Writes SetupInfo, GW / 2, GH - 130, 16, argb(255, 32, 32, 32), align:=StringAlignmentCenter
    Page.Writes format(Int(Progress * 1000) / 10, "0.0") & "%", GW / 2, GH - 100, 16, argb(255, 32, 32, 32), align:=StringAlignmentCenter, style:=FontStyleBold
End Sub
Public Sub DrawStep1()
    Dim Points(), y As Long
   
   
    Page.Writes "安装位置", 20, 60, 18, argb(255, 32, 32, 32), style:=FontStyleBold
    Page.Writes "该程序将被安装在您的计算机的下列位置。", 20, 90, 18, argb(255, 64, 64, 64)
    
    Page.Writes "您已选择信任该开发者：'" & SPackage.Maker & "'", 20, 440, 18, argb(255, 35, 170, 242), style:=FontStyleBold
    
    If Page.ShowEdit(SSetupPath, 0, 20, 150, GW - 40, 40, argb(255, 64, 64, 64), argb(255, 242, 242, 242), argb(255, 232, 232, 232), argb(255, 35, 170, 242), size:=18) = 4 Then
        On Error GoTo ErrHandler
        CreateFolder SSetupPath
ErrHandler:
        If Err.Number <> 0 Then SSetupPath = "C:\Program Files\" & SPackage.GameName: MsgBox "请使用其他目录。", 48
    End If
    
    If Page.ShowColorButton(0, GW - 80 - 20, 200, 80, 40, "浏览", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), size:=18) = 3 Then
        SSetupPath = GetDirName
        SSetupPath = SSetupPath & IIf(Right(SSetupPath, 1) <> "\", "\", "") & SPackage.GameName
        On Error GoTo ErrHandler2
        CreateFolder SSetupPath
ErrHandler2:
        If Err.Number <> 0 Then SSetupPath = "C:\Program Files\" & SPackage.GameName: MsgBox "请使用其他目录。", 48
    End If
    
    Page.ShowColorCheckBox LnkSwitch, 20, GH - 210, 200, 20, "创建桌面快捷方式", argb(255, 198, 198, 198), argb(255, 35, 170, 242), argb(255, 32, 32, 32), 18

    If Page.ShowColorButton(0, GW / 2 - 70, GH - 120, 140, 40, "确认", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=18) = 3 Then
        SetupInfo = "正在准备 ..."
        Step = 2
        ECore.NewTransform transDarkTo, 1000
        Dim Ret As String
        Ret = SetupPack
        If Ret <> "" Then
            'MsgBox "安装失败，请联系开发者，以下是错误信息：" & vbCrLf & Ret, 16, MainWindow.Caption
        End If
        ECore.NewTransform transDarkTo, 1000
        Step = 3
    End If
End Sub
Public Sub DrawStep0()
    Dim Points(), y As Long

    Page.Paint 1, 100, 122, 64, 64, argb(255, 92, 228, 220), pos:=posOnCenter
    Page.Paint 0, 0, 220, RGW, 400, argb(255, 242, 242, 242)
    
    Page.DrawImageEx "app.png", 100, 122, 64, 64, posOnCenter
    
    Page.Writes SPackage.GameName, 150, 93, 18, argb(255, 32, 32, 32)
    Page.Writes "开发者：" & SPackage.Maker, 150, 123, 18, argb(255, 64, 64, 64)
    
    Page.Writes "总计" & UBound(SPackage.Files) & "个文件将会被输出", 40, 240, 18, argb(255, 32, 32, 32)
    Page.Writes SPackage.GameDescribe, 40, 280, 18, argb(255, 64, 64, 64), RGW - 80, 100
    
    If Page.ShowColorButton(0, GW / 2 - 70, GH - 120, 140, 40, "安装", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=18) = 3 Then
        If MsgBox("确定要信任此程序的开发者：'" & SPackage.Maker & "'吗？" & vbCrLf & "并安装此程序到您的计算机。", vbYesNo + 48, "安装警告") = vbNo Then Exit Sub
        Step = 1
        ECore.NewTransform transDarkTo, 1000
    End If
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    CallByName Me, "DrawStep" & Step, VbMethod
    
    If Step <> 2 And Step <> 4 Then Call DrawTitle
End Sub
Public Sub DrawTitle()
    Page.Paint 0, 0, 0, RGW, 36, argb(255, 232, 232, 232)

    Dim m As Integer
    m = CheckMouse(GW - 36 - 8, 0, 36, 36)
    If Page.ShowColorButton(0, GW - 36 - 8, 0, 37, 36, "×", IIf(m = 0, argb(255, 27, 27, 27), argb(255, 255, 255, 255)), argb(255, 232, 232, 232), argb(255, 254, 84, 57), size:=18) = 2 Then
        Unload MainWindow
        End
    End If

    m = CheckMouse(GW - 72 - 8, 0, 36, 36)
    If Page.ShowColorButton(0, GW - 72 - 8, 0, 37, 36, "-", IIf(m = 0, argb(255, 27, 27, 27), argb(255, 255, 255, 255)), argb(255, 232, 232, 232), argb(255, 128, 128, 128), size:=18) = 2 Then
        MainWindow.WindowState = 1
        Mouse.State = 0
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
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

    '创建页面
    ECore.Add Page, "SetupPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

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
    Dim Page As GPage
    Public Progress As Single, SetupInfo As String, DoneMark As Boolean
    Dim DisplayProgress As Single, ChangeTime As Long
'======================================
Public Sub UpdateDone2()
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布
    
    If CmdMark = "Uninstall" Then GoTo uninpage
    
    If SetupErr = 0 Then
        Page.Writes "安装成功", GW / 2, GH / 2 - 40, 30, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    Else
        Page.Writes "安装已经完成，但是发生了一些错误。", GW / 2, 200, 20, argb(255, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes "错误代码：0x" & format(Hex(SetupErr), "000000"), GW / 2, 230, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes Error(SetupErr), GW / 2, 250, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        
        Page.Writes "请联系游戏开发者", GW / 2, 400, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter
    End If
    
    If Page.ShowColorButton(0, GW / 2 - 130 / 2, GH - 80, 130, 40, "OK", argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
        Unload MainWindow: End
    End If
    
    Exit Sub
    
uninpage:
    If SetupErr = 0 Then
        Page.Writes "卸载完成", GW / 2, GH / 2, 30, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    Else
        Page.Writes "卸载已经完成，但是发生了一些错误。", GW / 2, 200, 20, argb(255, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes "错误代码：0x" & format(Hex(SetupErr), "000000"), GW / 2, 230, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes Error(SetupErr), GW / 2, 250, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        
        Page.Writes "请联系开发者", GW / 2, 400, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter
    End If
End Sub
Public Sub UpdateDone()
    If SetupMode Then Call UpdateDone2: Exit Sub
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布
    
    If CmdMark = "Uninstall" Then GoTo uninpage
    
    If SetupErr = 0 Then
        Page.Writes "Emerald Builder 安装成功", GW / 2, 100, 30, argb(255, 27, 27, 27), align:=StringAlignmentCenter
        Page.Writes "在资源管理器中右键创建你的工程吧！", GW / 2, 140, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter, animation:=1
        
        Page.Paint 0, GW / 2, 160 + (GH - 160 - 80) / 2, GW + 1, 270, argb(255, 255, 255, 255), pos:=posOnCenter, animation:=2
        Page.DrawImage "tip.png", GW / 2, 160 + (GH - 160 - 80) / 2, Alpha:=1, pos:=posOnCenter, animation:=2
    Else
        Page.Writes "安装已经完成，但是发生了一些错误。", GW / 2, 200, 20, argb(255, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes "错误代码：0x" & format(Hex(SetupErr), "000000"), GW / 2, 230, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes Error(SetupErr), GW / 2, 250, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        
        Page.Writes "请联系QQ 1361778219 或者在Github上提交Issue", GW / 2, 400, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter
    End If
    
    If Page.ShowColorButton(0, GW / 2 - 130 / 2, GH - 80, 130, 40, "OK", argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
        Unload MainWindow: End
    End If
    
    Exit Sub
    
uninpage:
    If SetupErr = 0 Then
        Page.Writes "Emerald Builder 卸载完成", GW / 2, GH / 2, 30, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    Else
        Page.Writes "卸载已经完成，但是发生了一些错误。", GW / 2, 200, 20, argb(255, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes "错误代码：0x" & format(Hex(SetupErr), "000000"), GW / 2, 230, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        Page.Writes Error(SetupErr), GW / 2, 250, 16, argb(150, 254, 84, 57), align:=StringAlignmentCenter
        
        Page.Writes "请联系QQ 1361778219 或者在Github上提交Issue", GW / 2, 400, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter
    End If
End Sub
Public Sub Update()
    '绘制事件
    If DoneMark Then Call UpdateDone: Exit Sub
    
    If Progress <> DisplayProgress Then
        If GetTickCount - ChangeTime >= 10 Then
            ChangeTime = GetTickCount
            DisplayProgress = DisplayProgress + (Progress - DisplayProgress) / 30
        End If
        If DisplayProgress > Progress Or (Progress - DisplayProgress) < 0.001 Then DisplayProgress = Progress
    End If
        
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布
    
    Page.PaintArc GW / 2, GH / 2 - 20, 200, 200, 360, -90, argb(255, 230, 230, 230), 2, 1, posOnCenter
    Page.PaintArc GW / 2, GH / 2 - 20, 200, 200, DisplayProgress * 360, -90, argb(255, 95, 229, 221), 3, 1, posOnCenter
    Page.Writes Int(DisplayProgress * 100), GW / 2, GH / 2 - 52, 50, argb(255, 27, 27, 27), align:=StringAlignmentCenter, style:=FontStyleBold

    Page.Writes SetupInfo, GW / 2, GH - 110, 18, argb(255, 64, 64, 64), align:=StringAlignmentCenter
    
    If DisplayProgress >= 0.9999 Then Page.StartAnimation 1, 500: Page.StartAnimation 2, 700: ECore.NewTransform , 400: DoneMark = True
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub Enter()
    '页面进入事件
    DisplayProgress = 0: Progress = 0: SetupInfo = "准备中": DoneMark = False
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
    Page.Res.NewImages App.path & "\assets", 0.7
    '创建页面
    ECore.Add Page, "SetupPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
        Page.NewAnimation False, "FadeIn", 2, 400
        Page.NewAnimation False, "FadeIn", 2, 400
    '===============================================
End Sub

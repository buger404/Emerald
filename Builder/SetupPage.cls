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
    Page.ShowColorCheckBox ShowLog, GW / 2 - 100, GH / 2 - 20, 200, 20, "Show installer report", argb(255, 198, 198, 198), argb(255, 100, 62, 255), argb(255, 96, 96, 96), 18

    If Page.ShowColorButton(2, GW / 2 - 70, GH - 150, 140, 40, "Finish", argb(255, 255, 255, 255), argb(255, 100, 62, 255), argb(255, 140, 102, 255), 40, size:=18) = 3 Then
        If ShowLog Then ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
        Unload MainWindow
        End
    End If
    'ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
End Sub
Public Sub DrawStep4()
    Page.ShowLoading GW / 2 - 100, GH / 2 - 160, 200, 200, 9, argb(255, 100, 62, 255), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
    Page.Writes SetupInfo, GW / 2, GH / 2 + 160, 24, argb(255, 64, 64, 64), align:=StringAlignmentCenter
    Page.Writes format(Int(Progress * 1000) / 10, "0.0") & "%", GW / 2, GH / 2 + 210, 18, argb(255, 96, 96, 96), align:=StringAlignmentCenter
End Sub
Public Sub DrawStep3()
    Page.ShowColorCheckBox LaunchNow, GW / 2 - 100, GH / 2 - 60, 200, 20, "Launch right now", argb(255, 198, 198, 198), argb(255, 100, 62, 255), argb(255, 96, 96, 96), 18
    Page.ShowColorCheckBox ShowLog, GW / 2 - 100, GH / 2 - 20, 200, 20, "Show installer report", argb(255, 198, 198, 198), argb(255, 100, 62, 255), argb(255, 96, 96, 96), 18

    If Page.ShowColorButton(2, GW / 2 - 70, GH - 150, 140, 40, "Finish", argb(255, 255, 255, 255), argb(255, 100, 62, 255), argb(255, 140, 102, 255), 40, size:=18) = 3 Then
        If LaunchNow Then ShellExecuteA 0, "open", SSetupPath & "\app.exe", "", "", SW_SHOW
        If ShowLog Then ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
        Unload MainWindow
        End
    End If
    'ShellExecuteA 0, "open", LogPath, "", "", SW_SHOW
End Sub
Public Sub DrawStep2()
    Page.ShowLoading GW / 2 - 100, GH / 2 - 160, 200, 200, 9, argb(255, 100, 62, 255), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
    Page.Writes SetupInfo, GW / 2, GH / 2 + 160, 24, argb(255, 64, 64, 64), align:=StringAlignmentCenter
    Page.Writes format(Int(Progress * 1000) / 10, "0.0") & "%", GW / 2, GH / 2 + 210, 18, argb(255, 96, 96, 96), align:=StringAlignmentCenter
End Sub
Public Sub DrawStep1()
    Dim Points(), y As Long
    ReDim Points(3)
    
    y = GH / 2 - 75
    
    Points(0) = 0: Points(1) = 0
    Points(2) = 0: Points(3) = y
    
    Dim Step2 As Long
    Step2 = Int(GW / 14)
    
    For X = 0 To (GW + Step2) Step Step2
        ReDim Preserve Points(UBound(Points) + 2)
        Points(UBound(Points) - 1) = X
        Points(UBound(Points)) = y + Sin(GetTickCount / 200 + X / 10) * (GH * 0.2 / 12)
    Next
    
    ReDim Preserve Points(UBound(Points) + 4)
    Points(UBound(Points) - 3) = GW: Points(UBound(Points) - 2) = y
    Points(UBound(Points) - 1) = GW: Points(UBound(Points)) = 0
    
    Page.PaintPolygon2 argb(255, 27, 27, 27), 0, Points(), True
    
    Page.Writes "Install Path", 100, GH / 2 + 50, 18, argb(255, 100, 62, 255)
    
    If Page.ShowEdit(SSetupPath, 2, 88, GH / 2 + 85, GW - 88 * 2, 40, argb(255, 96, 96, 96), argb(255, 255, 255, 255), argb(255, 242, 242, 242), argb(255, 100, 62, 255), 90, size:=18) = 4 Then
        On Error GoTo ErrHandler
        CreateFolder SSetupPath
ErrHandler:
        If Err.Number <> 0 Then SSetupPath = "C:\Program Files\" & SPackage.GameName: MsgBox "Please try another folder .", 48
    End If
    
    Page.ShowColorCheckBox LnkSwitch, 98, GH / 2 + 260, 200, 20, "Create a desktop shortcut", argb(255, 198, 198, 198), argb(255, 100, 62, 255), argb(255, 96, 96, 96), 18

    If Page.ShowColorButton(2, GW - 200, GH - 130, 140, 40, "Start", argb(255, 255, 255, 255), argb(255, 100, 62, 255), argb(255, 140, 102, 255), 40, size:=18) = 3 Then
        SetupInfo = "Begining ..."
        Step = 2
        ECore.NewTransform transDarkTo, 1000
        Dim ret As String
        ret = SetupPack
        If ret <> "" Then
            MsgBox "Fail to install ." & vbCrLf & ret, 16, MainWindow.Caption
        End If
        ECore.NewTransform transDarkTo, 1000
        Step = 3
    End If
End Sub
Public Sub DrawStep0()
    Dim Points(), y As Long
    ReDim Points(3)
    
    y = GH / 2 - 75
    
    Points(0) = 0: Points(1) = 0
    Points(2) = 0: Points(3) = y
    
    Dim Step2 As Long
    Step2 = Int(GW / 14)
    
    For X = 0 To (GW + Step2) Step Step2
        ReDim Preserve Points(UBound(Points) + 2)
        Points(UBound(Points) - 1) = X
        Points(UBound(Points)) = y + Sin(GetTickCount / 200 + X / 10) * (GH * 0.2 / 12)
    Next
    
    ReDim Preserve Points(UBound(Points) + 4)
    Points(UBound(Points) - 3) = GW: Points(UBound(Points) - 2) = y
    Points(UBound(Points) - 1) = GW: Points(UBound(Points)) = 0
    
    Page.PaintPolygon2 argb(255, 27, 27, 27), 0, Points(), True

    Page.Paint 1, GW / 2, GH / 2 - 80, 96, 96, argb(255, 92, 228, 220), pos:=posOnCenter
    Page.DrawImageEx "app.png", GW / 2, GH / 2 - 80, 96, 96, posOnCenter
    
    Page.Writes "Here's " & SPackage.GameName, GW / 2, GH / 2 + 10, 24, argb(255, 64, 64, 64), align:=StringAlignmentCenter
    Page.Writes "Made by " & SPackage.Maker, GW / 2, GH / 2 + 50, 18, argb(255, 96, 96, 96), align:=StringAlignmentCenter
    Page.Writes "(" & SPackage.GameDescribe & ")", GW / 2, GH / 2 + 80, 18, argb(255, 96, 96, 96), align:=StringAlignmentCenter
    
    If Page.ShowColorButton(2, GW / 2 - 70, GH - 150, 140, 40, "Install", argb(255, 255, 255, 255), argb(255, 100, 62, 255), argb(255, 140, 102, 255), 40, size:=18) = 3 Then
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
    Dim m As Integer
    m = CheckMouse(GW - 60 - 19, 48, 36, 36)
    If Page.ShowColorButton(1, GW - 60 - 19, 35, 36, 36, "×", IIf(m = 0, argb(255, 128, 128, 128), argb(255, 255, 255, 255)), argb(0, 254, 84, 57), argb(255, 254, 84, 57), size:=18) = 2 Then
        Unload MainWindow
        End
    End If

    m = CheckMouse(GW - 120 - 19, 48, 36, 36)
    If Page.ShowColorButton(1, GW - 120 - 19, 35, 36, 36, "-", IIf(m = 0, argb(255, 128, 128, 128), argb(255, 255, 255, 255)), argb(0, 128, 128, 128), argb(100, 128, 128, 128), size:=18) = 2 Then
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

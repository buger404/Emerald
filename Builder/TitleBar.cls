VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TitleBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim SX As Long, SY As Long
'======================================

Public Sub Update()
    '绘制事件
    
    Page.Clear 0 '清空画布
    
    m = CheckMouse(231, 8, 98, 94)
    If m >= 1 And Mouse.button = 1 Then
        If SX = -1 Then
            SX = MainWindow.Left: SY = MainWindow.top
        End If
        ReleaseCapture
        SendMessageA MainWindow.Hwnd, WM_SYSCOMMAND, SC_MOVE Or HTCAPTION, 0
        NowY = MainWindow.top: TargetY = MainWindow.top
        Mouse.State = 0
    End If
    If m = 2 Then
        If MainWindow.Left = SX And MainWindow.top = SY Then
            TargetY = Screen.Height + MainWindow.Height * 2
        End If
        SX = -1: Mouse.State = 0: Mouse.button = 0
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
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
    Page.TopPage = True '置顶页面
    '创建页面
    ECore.Add Page, "TitleBar"
    SX = -1
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

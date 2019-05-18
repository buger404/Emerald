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
'======================================

Public Sub Update()
    '绘制事件
    
    Page.Clear 0 '清空画布

    Page.Writes "Emerald Builder", 20, 20, Color:=argb(255, 128, 128, 128), size:=16
    
    Dim m As Integer
    m = CheckMouse(GW - 60, 0, 60, 30 + 16 / 0.75)
    If Page.ShowColorButton(0, GW - 60, 0, 61, 30 + 16 / 0.75, "×", IIf(m = 0, argb(255, 128, 128, 128), argb(255, 255, 255, 255)), argb(0, 254, 84, 57), argb(255, 254, 84, 57), size:=18) = 2 Then
        Unload MainWindow
        Exit Sub
    End If
    m = CheckMouse(GW - 120, 0, 60, 30 + 16 / 0.75)
    If Page.ShowColorButton(0, GW - 120, 0, 61, 30 + 16 / 0.75, "-", IIf(m = 0, argb(255, 128, 128, 128), argb(255, 255, 255, 255)), argb(0, 128, 128, 128), argb(100, 128, 128, 128), size:=18) = 2 Then
        MainWindow.WindowState = 1
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
    Page.Res.NewImages App.Path & "\assets"
    Page.TopPage = True '置顶页面
    '创建页面
    ECore.Add Page, "TitleBar"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

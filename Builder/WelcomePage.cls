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

Public Sub Update()
    '绘制事件
    
    If InstalledPath <> "" Then GoTo OtherPage
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布

    Page.Writes "Welcome to Emerald Builder", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
    Page.Writes "Let's create our own games together .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
    
    Page.ShowColorButton 0, GW / 2 - 130 / 2, GH - 150, 130, 40, "Setup", argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16

    Exit Sub
    
OtherPage:
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
    '创建页面
    ECore.Add Page, "WelcomePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
        Page.NewAnimation False, "FadeIn", 2, 500
        Page.NewAnimation False, "FadeIn", 2, 500
    '===============================================
End Sub

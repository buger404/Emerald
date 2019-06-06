VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaitPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim pos As Single, degree As Single
    Dim first As Boolean
'======================================

Public Sub Update()
    '绘制事件
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布
    
    If Not first Then
        Page.Writes "请稍候", GW / 2, GH / 2, 28, argb(150, 27, 27, 27), align:=StringAlignmentCenter
        first = True
        Exit Sub
    End If
    
    Page.PaintArc GW / 2, GH / 2, 200, 200, 120, (GetTickCount Mod 3000) / 3000 * 360, argb(255, 232, 76, 61), 13, 1, posOnCenter
    Page.PaintArc GW / 2, GH / 2, 200, 200, 120, (GetTickCount Mod 1800) / 1800 * 360 + 90, argb(255, 240, 127, 34), 14, 1, posOnCenter
    Page.PaintArc GW / 2, GH / 2, 200, 200, 120, (GetTickCount Mod 1200) / 1200 * 360 + 90, argb(255, 242, 195, 17), 15, 1, posOnCenter

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

    '创建页面
    ECore.Add Page, "WaitPage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

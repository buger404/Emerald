VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DialogPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Public Key As Integer, Title As String, Context As String
    Dim buttons()
'======================================

Public Sub Update()
    '绘制事件
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布

    Page.Paint 0, GW / 2, GH / 2, 600, 250, argb(128, 255, 255, 255), pos:=posOnCenter
    Page.Writes Title, GW / 2 - 600 / 2 + 20, GH / 2 - 250 / 2 + 20, 30, argb(255, 27, 27, 27), style:=FontStyleBold
    Page.Writes Context, GW / 2 - 600 / 2 + 23, GH / 2 - 250 / 2 + 70, 18, argb(150, 27, 27, 27)
    
    For I = 0 To UBound(buttons)
        If Page.ShowColorButton(0, GW / 2 - 600 / 2 + 600 - 120 * (UBound(buttons) + 1) + I * 120, GH / 2 - 250 / 2 + 250 - 40 - 20, 100, 40, CStr(buttons(I)), argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
            Key = I + 1
        End If
    Next
    
    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
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
Public Sub NewDialog(t As String, C As String, btn())
    buttons = btn
    Title = t
    Context = C
    Key = 0
    ECore.NewTransform transFadeIn, 700, "DialogPage"
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源

    '创建页面
    ECore.Add Page, "DialogPage"
    
    '===============================================
    '   动画
    
    '===============================================
End Sub

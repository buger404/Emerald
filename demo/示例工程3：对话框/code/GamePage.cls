VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim DX As Long, DY As Long
'======================================

Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Page.DrawImage "background.png", 0, 0
    'Page.Paint 0, 0, 0, RGW + 1, RGH + 1, argb(255, 255, 255, 255)
    
    If Page.ShowColorButton(0, 115 - 22, 119, 115, 40, "测试1", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.SimpleMsg("这是一个有多个按钮的模糊背景信息框", "测试1", StrArray("好的好的", "知道了", "测试成功"))
    End If
    
    If Page.ShowColorButton(0, 287 - 22, 119, 115, 40, "测试2", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.SimpleMsg("这是一个只有单个按钮的模糊背景信息框", "测试2", StrArray("好的好的"))
    End If
    
    If Page.ShowColorButton(0, 287 + 172 - 22, 119, 115, 40, "测试3", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.SimpleMsg("这是一个自定义大小的模糊背景信息框", "测试3", StrArray("好的好的"), RGH * 0.5, RGH * 0.4)
    End If
    
    If Page.ShowColorButton(0, 287 + 172 * 2 - 22, 119, 115, 40, "测试4", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.SimpleMsg("这是一个不使用模糊背景的信息框", "测试4", StrArray("好的好的"), UseBlur:=False)
    End If
    
    If Page.ShowColorButton(0, 115 - 22, 119 + 80, 115, 40, "测试5", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.SimpleMsg("这是一个自定义模糊背景模糊半径的信息框", "测试5", StrArray("按钮1号", "非常非常非常非常长的按钮2", "按钮3"), Radius:=170)
    End If
    
    If Page.ShowColorButton(0, 287 - 22, 119 + 80, 115, 40, "测试6", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.ScreenMsg("这是一个简单的仅仅在屏幕提醒一句话的对话框")
    End If
    
    If Page.ShowColorButton(0, 287 + 172 - 22, 119 + 80, 115, 40, "测试7", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18) = mMouseUp Then
        Debug.Print Now, ECore.ScreenMsg("这是一个简单的仅仅在屏幕提醒一句话的对话框" & vbCrLf & "但是不使用模糊背景", False)
    End If
    
    Page.ShowColorButton 0, 287 + 172 * 2 - 22, 119 + 80, 115, 40, "测试8", argb(200, 0, 0, 0), argb(120, 0, 183, 195), argb(220, 0, 183, 195), size:=18
    ECore.TipMsg "测试8哟，这是一个类似Tooltip的东西~", 0
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
    Page.Res.NewImages App.path & "\assets", GW, GH
    '创建页面
    ECore.Add Page, "GamePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

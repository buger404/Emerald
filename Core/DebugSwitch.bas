Attribute VB_Name = "DebugSwitch"
'   Emerald 设置项

'======================================================
'   该设置已经迁移
'   相关设置请转到Builder中的“设置”
'======================================================


'==============================================================================
'   版本更新注意事项
'==============================================================================
'   1.鼠标点击修复
'      请在你的游戏窗口找到如下代码
'       Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'           If Mouse.State = 0 Then UpdateMouse X, Y, 0, button
'       End Sub
'      *****将其修改为：
'       Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
'           If Mouse.State = 0 Then
'               UpdateMouse X, Y, 0, button
'           Else
'               Mouse.X = X: Mouse.Y = Y
'           End If
'       End Sub
'   2.画布清空机制修改
'     请在你的绘图过程加入：
'       Page.Clear
'==============================================================================
'   1.资源加载的改变
'     请从Page.NewImages迁移到Page.Res.NewImages
'==============================================================================
'   1.加载代码的改变
'     由于开场LOGO的加入，
'     请把你设置主页面和开启绘制Timer的代码转移到创建页面之前，并加上一行“Me.Show”
'   *该部分您可以参照Emerald提供的模板
'   2.Emerald初始化的改变
'     您输入到Emerald的窗口大小，将会由Emerald重新设置一次。
'==============================================================================




'======================================================
'   警告：不要修改下列代码
    Public DebugMode As Integer, DisableLOGO As Integer, HideLOGO As Integer, UpdateCheckInterval As Long, UpdateTimeOut As Long
    Public Debug_focus As Boolean, Debug_pos As Boolean
'======================================================

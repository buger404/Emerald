Attribute VB_Name = "DebugSwitch"
'   Emerald 设置项

'======================================================
'   是否开启Debug
    Public Const DebugMode As Boolean = False
'   禁用开场LOGO
    Public Const DisableLOGO As Boolean = True
'   是否跳过多余的开场LOGO（如果资源已经加载完毕）
    Public Const HideLOGO As Boolean = False
'   检查更新间隔时长（天）
    Public Const UpdateCheckInterval As Long = 1
'   更新检查超时时间（毫秒）
    Public Const UpdateTimeOut As Long = 2000
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
'   Warning: please don't edit the following code .
    Public Debug_focus As Boolean
'======================================================

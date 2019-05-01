Attribute VB_Name = "DebugSwitch"
'======================================================
'   是否开启Debug
    Public Const DebugMode As Boolean = False
'   是否禁用开场LOGO（如果资源已经加载完毕）
    Public Const HideLOGO As Boolean = False
'   检查更新间隔时长（天）
    Public Const UpdateCheckInterval As Long = 1
'======================================================


'==============================================================================
'   版本更新注意事项
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

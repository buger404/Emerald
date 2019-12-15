VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Emerald Builder"
   ClientHeight    =   6672
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   9660
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   556
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
Private Sub DrawTimer_Timer()
    '绘制
    If EC.ActivePage = "" Then Exit Sub
    EC.Display
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    '初始化Emerald
    StartEmerald Me.Hwnd, 991, 754
    'ScaleGame 0.7, ScaleDefault

    DebugSwitch.HideLOGO = 1
    DebugSwitch.DisableLOGO = 1
    
    '创建字体
    Set EF = New GFont
    If PackPos = -1 Then
        EF.AddFont App.path & "\Builder.UI.otf"
        EF.MakeFont "Abadi MT Extra Light"
        'EF.MakeFont "微软雅黑"
    Else
        EF.MakeFont "微软雅黑"
    End If
    
    '创建页面管理器
    Set EC = New GMan
    If PackPos = -1 Then EC.Layered False
    
    '创建存档（可选）
    Set ESave = New GSaving
    ESave.Create "Emerald.Core"
    ESave.AutoSave = True
    
    '创建音乐列表
    Set MusicList = New GMusicList
    MusicList.Create App.path & "\music"

    '在此处初始化你的页面
    If PackPos = -1 Then
        Set WelcomePage = New WelcomePage
        Set ToNewPage = New ToNewPage
        Set TitleBar = New TitleBar
    Else
        Set SetupPage = New SetupPage
    End If

    ECore.FreezeMode = True

    '设置活动页面
    If PackPos = -1 Then EC.ActivePage = "WelcomePage"
    
    DrawTimer.Enabled = True
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
    If CmdMark = "Uninstall" Then
        Open VBA.Environ("temp") & "\copyemr.cmd" For Output As #1
        Print #1, "@echo off"
        Print #1, "echo 卸载程序正在清除残留文件 , Emerald Builder 版本号: " & Version
        Print #1, "echo 正在清理残留文件 ..."
        Print #1, "ping localhost -n 5 > nul"
        Print #1, "rd /s /q """ & App.path & """"
        Close #1
        ShellExecuteA 0, "open", VBA.Environ("temp") & "\copyemr.cmd", "", "", SW_SHOW
    End If
End Sub

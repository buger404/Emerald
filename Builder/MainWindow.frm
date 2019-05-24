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
    Dim oShadow As New aShadow
'==================================================
Private Sub DrawTimer_Timer()
    '绘制
    EC.Display
End Sub

Private Sub Form_Load()
    '初始化Emerald
    StartEmerald Me.Hwnd, 805, 556
    '创建字体
    MakeFont "微软雅黑"
    '创建页面管理器
    Set EC = New GMan
    EC.Layered True
    
    '创建存档（可选）
    Set ESave = New GSaving
    ESave.Create "Emerald.builder", "Emerald.builder"
    
    '创建音乐列表
    Set MusicList = New GMusicList
    MusicList.Create App.Path & "\music"

    '开始显示
    With oShadow
        If .Shadow(Me) Then
            .Depth = 20
            .Transparency = 16
        End If
    End With
    
    '在此处初始化你的页面
    Set WelcomePage = New WelcomePage
    Set SetupPage = New SetupPage
    Set WaitPage = New WaitPage
    Set DialogPage = New DialogPage
    Set UpdatePage = New UpdatePage
    
    Set TitleBar = New TitleBar

    '设置活动页面
    EC.ActivePage = "WaitPage"
    
    DrawTimer.Enabled = True
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    If Mouse.state = 0 Then
        UpdateMouse x, y, 0, button
    Else
        Mouse.x = x: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    '发送鼠标信息
    UpdateMouse x, y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
End Sub

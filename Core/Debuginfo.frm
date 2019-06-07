VERSION 5.00
Begin VB.Form Debuginfo 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Manager"
   ClientHeight    =   5748
   ClientLeft      =   36
   ClientTop       =   372
   ClientWidth     =   6084
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000005&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Reporter 
      Interval        =   20
      Left            =   5496
      Top             =   5208
   End
End
Attribute VB_Name = "Debuginfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Emerald 相关代码

Private Sub Reporter_Timer()
    If Not Me.Visible Then Exit Sub

    On Error Resume Next
    
    Me.Cls
    Me.CurrentX = 30: Me.CurrentY = 30
    
    Me.ForeColor = RGB(0, 0, 0)
    Me.CurrentX = 30
    Print "工程名：" & App.Title & vbCrLf
    Me.ForeColor = RGB(0, 176, 240)
    Me.CurrentX = 30
    Print "鼠标状态：" & Mouse.state & "(" & Mouse.button & ")  (" & Mouse.x & "," & Mouse.y & ")"
    
    Me.ForeColor = RGB(113, 119, 66)
    Me.CurrentX = 30
    Print "存档状态：" & IIf(Not ESave Is Nothing, "已创建", "未创建")
                                
    If Not ESave Is Nothing Then Me.CurrentX = 30: Print "权限：" & ESave.sToken & "，数据个数：" & ESave.Count
    
    Me.ForeColor = RGB(0, 0, 0)
    
    Me.CurrentX = 30
    Print vbCrLf
    Me.CurrentX = 30
    Print "当前活动页面：" & ECore.ActivePage
    Me.CurrentX = 30
    Print "FPS：" & FPS
    Me.CurrentX = 30
    Print "每帧耗时：" & Int(FPSct / FPS) & "ms"
    Me.CurrentX = 30
    Print "估测极限fps：" & Int(1000 / Int(FPSct / FPS))
    
    Me.ForeColor = RGB(255, 0, 0)
    
    Me.CurrentX = 30
    Print vbCrLf
    Me.CurrentX = 30
    Print "注意事项"
    
    If Abs(FPSctt - 1000) > 60 Then Me.CurrentX = 30: Print "似乎你正在使用Timer绘图。"

End Sub

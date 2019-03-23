VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Page As GPage, sy As Long
Public Sub Update()
    Page.DrawImage "404.png", 20, 20, alpha:=1, animation:=1
    Page.DrawImage "1.png", GW - 160, 20, alpha:=1, animation:=1
    If CheckMouse2 = 3 Then MsgBox "冰棍被鼠标打了": ECore.NewTransform 10, 1000
    Page.Writes "冰棍", GW - 80, 170, align:=StringAlignmentCenter, Color:=argb(255, 0, 0, 0), animation:=1
    
    '卷轴！
    Page.Paint 0, 20, 230, 200 + 1, 300 + 1, argb(255, 180, 180, 180), style:=1, animation:=1
    
    Page.StartScroll 20, 230, 200, 300, 0, sy
        Page.DrawImage "404.png", 0, 0, alpha:=0.5, animation:=1
        Page.DrawImage "test.gif", 18, 220, alpha:=1, animation:=1
        Page.Writes "ScrollY: " & sy, 0, 237, 16, argb(255, 255, 0, 0), 200, 20, align:=StringAlignmentCenter, style:=FontStyleBold, animation:=1
    Page.EndScroll
    
    If IsKeyPress(VK_UP) Then sy = sy + 10
    If IsKeyPress(VK_DOWN) Then sy = sy - 10
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    sy = sy + 10 * Depth * Direction
End Sub
Public Sub Enter()
    Page.StartAnimation 1
    Page.StartAnimation 2
End Sub
Public Sub Leave()

End Sub
Public Sub Gotfocus()
    Page.StartAnimation 1
    Page.StartAnimation 2
End Sub
Public Sub Lostfocus()
    ECore.NewTransform 11, 1000
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.NewImages App.path & "\assets"
    
    ECore.Add Page, "TestPage"
    
    '添加动画
    Page.NewAnimation False, "FlyInFromLeftFade", 2, 500
    Page.NewAnimation False, "FlyInFromTopFade", 2, 500
End Sub

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

    If (Not Page.AnimationDone(3)) Or (Not Page.AnimationDone(4)) Then
        Page.Paint 0, 0, 0, GW, GH, argb(255, 32, 32, 32), animation:=IIf(Page.AnimationDone(3), 4, 3)
        Page.Writes "»¶Ó­Æ·³¢ÐÂÏÊ±ù¹÷", GW / 2, GH / 2 - 20, 40, argb(255, 255, 255, 255), animation:=IIf(Page.AnimationDone(3), 4, 3), align:=StringAlignmentCenter
        Exit Sub
    End If

    Page.DrawImage "404.png", 20, 20, Alpha:=1, animation:=1
    Page.DrawImage "1.png", GW - 160, 20, Alpha:=1, animation:=1
    If CheckMouse2 = mMouseUp Then StartScreenDialog 600, 400, SysPage
    Page.Writes "±ù¹÷", GW - 80, 170, align:=StringAlignmentCenter, Color:=argb(255, 0, 0, 0), animation:=1
    
    '¾íÖá£¡
    Page.Paint 0, 20, 230, 200 + 1, 300 + 1, argb(255, 180, 180, 180), style:=1, animation:=1
    
    Page.StartScroll 20, 230, 200, 300, 0, sy
        Page.DrawImage "404.png", 0, 0, Alpha:=0.5, animation:=1
        Page.DrawImage "test.gif", 18, 220, Alpha:=1, animation:=1
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
    Page.StartAnimation 3
End Sub
Public Sub Leave()

End Sub
Public Sub Gotfocus()

End Sub
Public Sub Lostfocus()

End Sub
Public Sub AnimationDone(id As Integer)
    If id = 3 Then Page.StartAnimation 4, 2000
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.NewImages App.path & "\assets"
    
    ECore.Add Page, "TestPage"
    
    'Ìí¼Ó¶¯»­
    Page.NewAnimation False, "FlyInFromLeftFade", 2, 500
    Page.NewAnimation False, "FlyInFromTopFade", 2, 500
    Page.NewAnimation False, "FadeIn", 2, 500
    Page.NewAnimation False, "FadeOut", 3, 500

End Sub

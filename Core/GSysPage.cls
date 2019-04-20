VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSysPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Page As GPage, Index As Integer, Win As Object
Public Key As Integer
Public Sub Update()
    Select Case Index
        Case 0
        Call SaveDialog
    End Select
End Sub
Public Sub SaveDialog()
    Page.Paint 0, 0, 0, 600, 76, argb(255, 242, 242, 242)
    Page.DrawImage "save.png", 40, 40, Alpha:=1
    Page.Writes "你希望该应用在以下位置保存数据吗？" & vbCrLf & ESave.path, 40, 130, 18, argb(255, 64, 64, 64)
    
    Page.Writes "按下对应的键进行操作，如果不允许，该应用可能永远保持初始状态。", 40, 400 - 110, 16, argb(255, 96, 96, 96)
    
    Page.Paint 0, 600 - 140 * 2 + 20, 400 - 60, 100, 30, argb(255, 95, 229, 221)
    Page.Writes "允许（Y）", 600 - 140 * 2 + 20, 400 - 57, 16, argb(255, 255, 255, 255), 110, 30, StringAlignmentCenter
    If IsKeyPress(VK_Y) Then Key = 1: Win.Accept
    
    Page.Paint 0, 600 - 140 * 1 + 10, 400 - 60, 100, 30, argb(255, 180, 180, 180)
    Page.Writes "拒绝（N）", 600 - 140 * 1 + 10, 400 - 57, 16, argb(255, 48, 48, 48), 110, 30, StringAlignmentCenter
    If IsKeyPress(VK_N) Then Key = 0: Win.Accept
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)

End Sub
Public Sub Enter()

End Sub
Public Sub Leave()

End Sub
Public Sub Gotfocus()

End Sub
Public Sub Lostfocus()

End Sub
Public Sub AnimationDone(id As Integer)

End Sub
Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.NewImages App.path & "\assets\debug"
End Sub

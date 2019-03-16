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
Dim Page As GPage, SY As Long
Public Sub Update()
    Page.DrawImage "404.png", 20, 20, alpha:=1
    Page.DrawImage "1.png", GW - 160, 20, alpha:=1
    If CheckMouse2 = 3 Then MsgBox "±ù¹÷±»Êó±ê´òÁË": ECore.NewTransform 10, 1000
    Page.Writes "±ù¹÷", GW - 80, 170, align:=StringAlignmentCenter, Color:=argb(255, 0, 0, 0)
    
    '¾íÖá£¡
    Page.Paint 0, 20, 230, 200 + 1, 300 + 1, argb(255, 180, 180, 180), style:=1
    
    Page.StartScroll 20, 230, 200, 300, 0, SY
        Page.DrawImage "404.png", 0, 0, alpha:=0.5
        Page.DrawImage "test.gif", 18, 220, alpha:=1
        Page.Writes "ScrollY: " & SY, 0, 237, 16, argb(255, 255, 0, 0), 200, 20, align:=StringAlignmentCenter, style:=FontStyleBold
    Page.EndScroll
    
    If IsKeyPress(VK_UP) Then SY = SY + 10
    If IsKeyPress(VK_DOWN) Then SY = SY - 10
End Sub

Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.NewImages App.path & "\assets"
    
    ECore.Add Page, "TestPage"
End Sub

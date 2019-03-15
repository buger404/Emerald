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
Dim Page As GPage
Public Sub Update()
    Page.DrawImage "404.png", 20, 20, alpha:=1
    Page.DrawImage "1.png", GW - 160, 20, alpha:=1
    If CheckMouse2 = 3 Then MsgBox "冰棍被鼠标打了"
    Page.Writes "大爱冰棍（雾）", GW - 80, 170, align:=StringAlignmentCenter
End Sub

Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.NewImages App.path & "\assets"
    
    ECore.Add Page, "TestPage"
End Sub

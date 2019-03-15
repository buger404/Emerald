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
    If CheckMouse2 = 3 Then MsgBox "±ù¹÷±»Êó±ê´òÁË"
    Page.Writes "´ó°®±ù¹÷£¨Îí£©", GW - 80, 170, align:=StringAlignmentCenter, Color:=argb(255, 0, 0, 0)
    
    '¾íÖá£¡
    Page.StartScroll 20, 230, 200, 300, 0, -140
        Page.DrawImage "404.png", 0, 0, alpha:=0.5
        Page.DrawImage "test.gif", 18, 220, alpha:=1
        Page.Writes "±ù¹÷ºÃÀ«ÅÂ", 0, 237, 16, argb(255, 255, 0, 0), 200, 20, align:=StringAlignmentCenter, style:=FontStyleBold
    Page.EndScroll
End Sub

Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.NewImages App.path & "\assets"
    
    ECore.Add Page, "TestPage"
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Page As GPage, GW As Long, GH As Long
Public Sub Update()
    
    Page.Writes FPS, 15, 14, 48, argb(205, 95, 229, 221)
    
    Page.DrawImage "mouse" & Mouse.button & ".png", GW - 170, 5, Alpha:=1
    Page.DrawImage "menu.png", GW - 90, 5, Alpha:=1
    
    Page.DrawImage "focus.png", GW - 250, 5, Alpha:=1
    If Debug_focus Then Page.Paint 0, GW - 250, 0, 80, GH, argb(50, 95, 229, 221)
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

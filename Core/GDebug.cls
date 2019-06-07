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
'Emerald Ïà¹Ø´úÂë

Public Page As GPage, GW As Long, GH As Long
Public Sub Update()
    On Error Resume Next
    
    Page.Writes FPS, 15, 14, 34, argb(205, 182, 174, 173)
    
    Page.DrawImage "mouse" & Mouse.button & ".png", GW - 138, 78 / 2 - 64 / 2, alpha:=1
    Page.DrawImage "menu.png", GW - 74, 78 / 2 - 64 / 2, alpha:=1

    Page.DrawImage "focus.png", GW - 202, 78 / 2 - 64 / 2, alpha:=1
    If Debug_focus Then Page.Paint 0, GW - 202, 0, 64, 79, argb(50, 128, 128, 128)
    
    Page.DrawImage "pos.png", GW - 202 - 64, 78 / 2 - 64 / 2, alpha:=1
    If Debug_pos Then Page.Paint 0, GW - 202 - 64, 0, 64, 79, argb(50, 128, 128, 128)
    
    Page.DrawImage "data.png", GW - 202 - 128, 78 / 2 - 64 / 2, alpha:=1
    If Debug_data Then Page.Paint 0, GW - 202 - 128, 0, 64, 79, argb(50, 128, 128, 128)
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

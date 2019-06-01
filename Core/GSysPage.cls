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
'Emerald 相关代码

Public Page As GPage, index As Integer, Win As Object
Public Key As Integer
Public OpenTime As Long, DoneMark As Boolean, DoneStep As Integer
Dim ChangeTime As Long, LastLoadPro As Single, LastLoadPro2 As Single
Public Sub Update()
    Page.Clear
    Select Case index
        Case 0
        Call SaveDialog
        Case 1
        Call LOGODialog
    End Select
End Sub
Public Sub LOGODialog()
    If DisableLOGO Then
        DoneMark = True
        Exit Sub
    End If
    
    Dim w As Long, h As Long
    w = Page.Res.ImgSize("logo.png", imgGetWidth)
    h = Page.Res.ImgSize("logo.png", imgGetHeight)
    If w = 0 Then Exit Sub
    
    Page.Clear argb(255, 16, 16, 35)
    
    If GetTickCount - OpenTime >= 2000 Then
        If PreLoadCount <> LoadedCount Then
            Page.DrawImageEx "logo.png", GW / 2, GH * 0.85 / 2, Int(GW * 0.45), Int((GW * 0.45) / w * h), 1
        End If
    Else
        Page.DrawImageEx "logo.png", GW / 2, GH / 2, Int(GW * 0.45), Int((GW * 0.45) / w * h), 1
    End If
    
    Dim pro As Single
    
    If (GetTickCount - OpenTime <= 1000) And (Not HideLOGO) Then
        pro = Cubic((GetTickCount - OpenTime) / 1000, 0, 0, 0, 1)
        Page.Paint 0, 0, 0, GW, GH, argb(255 - pro * 255, 16, 16, 35)
    End If
    
    If HideLOGO And PreLoadCount = LoadedCount Then
        ECore.NewTransform 1, 1000
        DoneMark = True
    End If
    
    If GetTickCount - OpenTime >= 2000 Then
        If PreLoadCount = LoadedCount And ECore.ActivePage <> "" Then
            ECore.NewTransform 1, 1000
            DoneMark = True
        Else
            If DoneStep = 0 Then
                DoneStep = 1
                ECore.NewTransform 0, 1000
            End If
            pro = LoadedCount / PreLoadCount
            If pro > 1 Then pro = 1
            If LastLoadPro <> pro Then LastLoadPro2 = LastLoadPro: LastLoadPro = pro: ChangeTime = GetTickCount
            If GetTickCount - ChangeTime <= 500 Then
                pro = LastLoadPro2 + (pro - LastLoadPro2) * (GetTickCount - ChangeTime) / 500
            End If
            Page.Paint 2, GW / 2 - GW * 0.5 / 2, GH * 0.8, GW * 0.5, 14, argb(255, 127, 127, 127), 14
            Page.Paint 2, GW / 2 - GW * 0.5 / 2, GH * 0.8, GW * 0.5 * pro, 14, argb(255, 255, 255, 255), 14
        End If
    End If
End Sub
Public Sub SaveDialog()
    Page.Paint 0, 0, 0, 600, 76, argb(255, 242, 242, 242)
    Page.DrawImage "save.png", 40, 40, Alpha:=1
    Page.Writes "你希望该应用在以下位置保存数据吗？" & vbCrLf & ESave.path, 40, 130, 18, argb(255, 64, 64, 64)
    
    Page.Writes "*如果不允许，该应用可能永远保持初始状态。", 40, 400 - 110, 16, argb(255, 96, 96, 96)
    
    Dim MouseT As Integer
    MouseT = CheckMouse(600 - 140 * 2 + 20, 400 - 60, 100, 30)
    
    Page.Paint 0, 600 - 140 * 2 + 20, 400 - 60, 100, 30, IIf(MouseT <> 0, argb(255, 247, 181, 99), argb(255, 223, 165, 92))
    Page.Writes "允许（Y）", 600 - 140 * 2 + 20, 400 - 57, 16, argb(255, 255, 255, 255), 110, 30, StringAlignmentCenter
    If IsKeyPress(VK_Y) Or (MouseT > 1) Then Key = 1: Win.Accept
    
    MouseT = CheckMouse(600 - 140 * 1 + 10, 400 - 60, 100, 30)
    
    Page.Paint 0, 600 - 140 * 1 + 10, 400 - 60, 100, 30, IIf(MouseT <> 0, argb(255, 180, 180, 180), argb(255, 200, 200, 200))
    Page.Writes "拒绝（N）", 600 - 140 * 1 + 10, 400 - 57, 16, argb(255, 48, 48, 48), 110, 30, StringAlignmentCenter
    If IsKeyPress(VK_N) Or (MouseT > 1) Then Key = 0: Win.Accept
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
    Page.Res.NewImages App.path & "\assets\debug"
End Sub

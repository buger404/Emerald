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
Public BDC As Long
Public Key As Integer
Public OpenTime As Long, DoneMark As Boolean, DoneStep As Integer
Public DataPage As Integer
Dim ChangeTime As Long, LastLoadPro As Single, LastLoadPro2 As Single
Dim max2 As Long
Public Sub Update()
    Page.Clear
    Select Case index
        Case 0
        Call SaveDialog
        Case 1
        Call LOGODialog
        Case 3
        Call SaveManDialog
    End Select
End Sub
Public Sub SaveManDialog()
    Page.Clear argb(220, 242, 242, 249)
    
    Page.Paint 2, GW * 0.2, GH * 0.1, GW * 0.6, GH * 0.8, argb(255, 255, 255, 255), 30
    Page.Writes "存档管理", GW / 2, 83, 18, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    
    If ESave Is Nothing Then
        Page.Writes "此页面不可用，你没有创建存档。", GW / 2, GH / 2, 18, argb(150, 27, 27, 27), align:=StringAlignmentCenter
        Exit Sub
    End If
    
    Page.Writes "key " & IIf(ESave.Key <> "", ESave.Key, "[Default]"), GW / 2, 107, 14, argb(150, 27, 27, 27), align:=StringAlignmentCenter

    Dim temp As String, temp2 As String

    If Page.ShowColorButton(1, GW * 0.2 + GW * 0.6 - 24, GH * 0.1 - 24, 48, 48, "+", argb(255, 255, 255, 255), argb(130, 95, 229, 221), argb(255, 95, 229, 221), size:=24, style:=FontStyleBold) = 3 Then
        temp2 = InputBox("输入新的数据的名称", "新建数据")
        If temp2 = "" Then Exit Sub
        temp = InputBox("输入新的值：" & vbCrLf & temp2, "存档修改")
        If temp = "" Then Exit Sub
        ESave.PutData temp2, temp
    End If

    Dim max As Long, i As Integer
    max = Int(GH * 0.5 / 30)

    For i = DataPage + 1 To DataPage + max
        If i > ESave.Count Then Exit For
        Page.Paint 1, GW / 2 - GW * 0.3 + 20, 80 + 30 + (i - DataPage) * 30 + 2, 20, 20, argb(255, 0, 176, 240)
        If CheckMouse2 = mMouseUp Then
            temp = InputBox("输入新的值：" & vbCrLf & ESave.GetName(i), "存档修改", ESave.GetData(ESave.GetName(i)))
            ESave.PutData ESave.GetName(i), temp
        End If
        Page.Writes ESave.GetName(i), GW / 2 - GW * 0.3 + 50, 80 + 30 + (i - DataPage) * 30, 18, argb(150, 27, 27, 27), align:=StringAlignmentNear, style:=FontStyleBold
        Page.Writes ESave.GetData(ESave.GetName(i)), GW / 2 + GW * 0.3 - 30, 80 + 30 + (i - DataPage) * 30, 18, argb(150, 27, 27, 27), align:=StringAlignmentFar
    Next
    
    Page.Writes (Int(DataPage / max) + 1) & " / " & (Int(ESave.Count / max) + 1), GW / 2, GH - 120, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter
    If Page.ShowColorButton(2, GW / 2 - 80 - 20, GH - 120, 40, 22, "<", argb(150, 27, 27, 27), argb(255, 222, 222, 222), argb(130, 95, 229, 221), size:=16, Radius:=60) = 3 Then
        If DataPage > 0 Then DataPage = DataPage - max: ECore.NewTransform
    End If
    If Page.ShowColorButton(2, GW / 2 + 80 - 20, GH - 120, 40, 22, ">", argb(150, 27, 27, 27), argb(255, 222, 222, 222), argb(130, 95, 229, 221), size:=16, Radius:=60) = 3 Then
        If (Int(DataPage / max) + 1) < (Int(ESave.Count / max) + 1) Then DataPage = DataPage + max: ECore.NewTransform
    End If
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
    Page.DrawImage "save.png", 40, 40, alpha:=1
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
    
    BDC = CreateCDC(GW, GH)
End Sub

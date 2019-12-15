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
Dim LOGO As Long, bitmap As Long
Dim PXPitch As Long
Dim EnterTime As Long
Public Sub Update()
    Page.Clear
    Select Case index
        Case 0
        Call SaveDialog
        Case 1
        Call LOGODialog
        Case 3
        Call SaveManDialog
        Case 4
        Call SuggestionDialog
        Case 5
        Call PageDialog
    End Select
End Sub
Public Sub PageDialog()
    Page.Clear argb(0, 0, 0, 0)
    Page.Paint 0, 0, GH - 100, RGW + 1, 101, argb(80, 0, 0, 0)
    Dim I As Integer, Active As Boolean
    Dim acI As Integer
    For I = 1 To ECore.PageCount
        If ECore.ActivePage = ECore.PageName(I) Then acI = I: Exit For
    Next
    
    If IsKeyPress(VK_LEFT) Then PXPitch = PXPitch + 10
    If IsKeyPress(VK_RIGHT) Then PXPitch = PXPitch - 10
    
    If PXPitch < 0 Then PXPitch = 0
    If PXPitch > (ECore.PageCount - 1) * 170 Then PXPitch = (ECore.PageCount - 1) * 170
    
    For I = 1 To ECore.PageCount
        Active = (acI = I)
        Page.Paint 2, 30 + (I - 1) * 170 - PXPitch, GH - 73, 150, 40, IIf(Active, argb(255, 255, 84, 57), argb(180, 242, 242, 242)), 100
        If CheckMouse2 = mMouseUp And (Not Active) Then ECore.NewTransform IIf(I >= acI, transDarkTo, transDarkReturn), , ECore.PageName(I)
        If Not Active Then Page.Paint 2, 30 + (I - 1) * 170 - PXPitch, GH - 73, 150, 40, argb(255, 200, 200, 200), 100, style:=1
        Page.Writes ECore.PageName(I), 30 + (I - 1) * 170 + 150 / 2 - PXPitch, GH - 66, 18, IIf(Active, argb(255, 255, 255, 255), argb(255, 32, 32, 32)), align:=StringAlignmentCenter
    Next
End Sub
Public Sub SuggestionDialog()
    Page.Clear argb(0, 0, 0, 0)
    If IsKeyPress(VK_F12) Then
        If UBound(SGS) = 0 Then
            MsgBox "没有任何建议。", 48, "Emerald 建议中心"
        Else
            Dim Ret As String, sRet As String
            Ret = "共 " & UBound(SGS) & " 项建议。" & vbCrLf & vbCrLf
            For I = 1 To UBound(SGS)
                Select Case SGS(I).Deepth
                    Case 0: sRet = "(无关紧要)"
                    Case 1: sRet = "(需要留意)"
                    Case 2: sRet = "(*紧急)"
                    Case 3: sRet = "(*非常紧急*)"
                End Select
                Ret = Ret & sRet & "    " & SGS(I).Content & vbCrLf
            Next
            MsgBox Ret, 48, "Emerald 建议中心"
        End If
    End If
    If GetTickCount - SGTime <= 3000 And UBound(SGS) <> 0 Then
        Dim pro As Single
        
        If GetTickCount Mod 600 < 300 Then
            pro = 0.6 + (GetTickCount Mod 300) / 300 * 0.4
        Else
            pro = 1 - (GetTickCount Mod 300) / 300 * 0.4
        End If
        
        Page.Paint 2, GW / 2, GH / 2, 300, 40, argb(255 * pro, 255, 84, 57), 100, pos:=posOnCenter
        Page.Writes "共 " & UBound(SGS) & " 项建议 (F12)", GW / 2, GH / 2 - 8 / 0.75, 16, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    End If
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

    Dim max As Long, I As Integer
    max = Int(GH * 0.5 / 30)

    For I = DataPage + 1 To DataPage + max
        If I > ESave.Count Then Exit For
        Page.Paint 1, GW / 2 - GW * 0.3 + 20, 80 + 30 + (I - DataPage) * 30 + 2, 20, 20, argb(255, 0, 176, 240)
        If CheckMouse2 = mMouseUp Then
            temp = InputBox("输入新的值：" & vbCrLf & ESave.GetName(I), "存档修改", ESave.GetData(ESave.GetName(I)))
            ESave.PutData ESave.GetName(I), temp
        End If
        Page.Writes ESave.GetName(I), GW / 2 - GW * 0.3 + 50, 80 + 30 + (I - DataPage) * 30, 18, argb(150, 27, 27, 27), align:=StringAlignmentNear, style:=FontStyleBold
        Page.Writes ESave.GetData(ESave.GetName(I)), GW / 2 + GW * 0.3 - 30, 80 + 30 + (I - DataPage) * 30, 18, argb(150, 27, 27, 27), align:=StringAlignmentFar
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

    If EnterTime = 0 Then EnterTime = GetTickCount
    Dim tpro As Single, tpro2 As Single
    tpro = (GetTickCount - EnterTime) / 1000
    If tpro > 1 Then tpro = 1
    tpro = Cubic(tpro, 0, 1, 1, 1)
    tpro2 = (GetTickCount - EnterTime - 1000) / 1000
    If tpro2 > 1 Then tpro2 = 1
    If tpro2 < 0 Then tpro2 = 0
    tpro2 = Cubic(tpro2, 0, 1, 1, 1)
    
    Page.Clear argb(255, 255, 255, 255)
        
    Dim pro As Single
    pro = (LoadedCount / PreLoadCount)
    
    Page.DrawImage "profile.png", GW / 2, GH / 2 - 90 * tpro, alpha:=1 * tpro, pos:=posOnCenter
    Page.DrawImage "elogo.png", GW / 2 - 115, GH - 45, alpha:=1 * tpro2, pos:=posOnCenter
    Page.Writes "Powered by Emerald", GW / 2 + 25, GH - 57, 18, argb(170 * tpro2, 27, 27, 27), align:=StringAlignmentCenter
<<<<<<< HEAD
    
    Page.ShowLoading GW / 2 - 24, GH / 2 + GH * 0.25 - 24, 48, 48, 4, argb(255 * tpro2, 100, 62, 255), argb(255 * tpro2, 0, 176, 240), argb(255 * tpro2, 64, 218, 208)
    
    Page.Writes App.ProductName, GW / 2, GH / 2 + 100 + (GH * 0.2 / 2 - 160) * tpro, 20, argb(255 * tpro, 27, 27, 27), align:=StringAlignmentCenter
    
    Page.Writes "正在启动 " & Int(pro * 100) & "%", GW / 2, GH / 2 + (GH * 0.2 / 2 + 17 - 40) * tpro2, 18, argb(200 * tpro2, 27, 27, 27), align:=StringAlignmentCenter
    
=======
    
    Page.ShowLoading GW / 2 - 24, GH / 2 + 100, 48, 48, 4, argb(255 * tpro2, 100, 62, 255), argb(255 * tpro2, 0, 176, 240), argb(255 * tpro2, 64, 218, 208)
    
    Page.Writes App.ProductName, GW / 2, GH / 2 + 100 + (GH * 0.2 / 2 - 160) * tpro, 20, argb(255 * tpro, 27, 27, 27), align:=StringAlignmentCenter
    
    Page.Writes "正在启动 " & Int(pro * 100) & "%", GW / 2, GH / 2 + (GH * 0.2 / 2 + 17 - 40) * tpro2, 18, argb(200 * tpro2, 27, 27, 27), align:=StringAlignmentCenter
    
>>>>>>> 7da8aea3f1f11a7fbe211537042b7146cb726c86
    If DisableLOGO Then
        ECore.NewTransform transToLeft, 1000
        DoneMark = True
        Exit Sub
    End If
    If PreLoadCount = LoadedCount And (GetTickCount - EnterTime) >= 3000 Then
        ECore.NewTransform transToLeft, 1000
        DoneMark = True
    End If
    
End Sub
Public Sub SaveDialog()
    Page.Paint 0, 0, 0, 600, 76, argb(255, 242, 242, 242)
    Page.DrawImage "save.png", 40, 40, alpha:=1
    Page.Writes "你希望该应用在以下位置保存数据吗？" & vbCrLf & ESave.Path, 40, 130, 18, argb(255, 64, 64, 64)
    
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
Public Sub AnimationDone(ID As Integer)

End Sub
Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    Page.Res.NewImages App.Path & "\assets\debug"
<<<<<<< HEAD
    If Dir(App.Path & "\assets\sets\profile.png") <> "" Then
        Page.Res.newImage App.Path & "\assets\sets\profile.png", 148, 148, "profile.png"
    Else
        Page.Res.newImage App.Path & "\assets\debug\icon.png", 148, 148, "profile.png"
    End If
=======
    Page.Res.newImage App.Path & "\assets\debug\icon.png", 148, 148, "profile.png"
>>>>>>> 7da8aea3f1f11a7fbe211537042b7146cb726c86
    Page.Res.newImage App.Path & "\assets\debug\logo.png", 64, 64, "elogo.png"
    
    Page.TopPage = True
    
    BDC = CreateCDC(GW, GH)
End Sub


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
Public MsgButton As Integer
Dim RectMode As Integer
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
        Case 6
        Call DesignDialog
        Case 7
        Call MsgDialog
    End Select
End Sub
Public Sub ScreenMsgDrawer()
    Dim upro As Single
    upro = (GetTickCount - MsgEnterTime) / 500
    If upro > 1 Then upro = 1
    upro = Cubic(upro, 0, 1, 1, 1)

    Dim m As MButtonState
    
    Page.Clear 0
    PaintDC BDC, Page.CDC, alpha:=1
    
    Page.Paint 0, 0, 0, RGW + 1, RGH + 1, argb(80, 0, 0, 0)
    
    Page.Writes MsgContent, RGW / 2, RGH / 2 + RGH / 2 * 0.2 * (1 - upro) - 10, 20, argb(255 * upro, 255, 255, 255), align:=StringAlignmentCenter, BorderColor:=argb(160 * upro, 16, 16, 16), BorderSize:=1, style:=FontStyleBold

    m = CheckMouse(0, 0, RGW, RGH)
    
    If m = mMouseUp Then Me.MsgButton = 1
End Sub
Public Sub TipMsgDrawer()
    Dim m As MButtonState
    
    Dim pro As Single, bpro As Single
    If MsgBtnFocusTime <> 0 Then
        pro = (GetTickCount - MsgBtnFocusTime) / IIf(MsgBtnFocus = -1, 500, 200)
        If pro > 1 Then pro = 1
        pro = Cubic(pro, 0, 1, 1, 1)
        If MsgBtnFocus = -1 Then
            bpro = pro
        Else
            bpro = 1 - pro
        End If
    End If
    
    Page.Clear 0
    PaintDC BDC, Page.CDC, MsgX + MsgW / 2 * (1 - bpro), MsgY, MsgX + MsgW / 2 * (1 - bpro), MsgY, MsgW * bpro, MsgH, bpro
    
    Dim c(4) As Byte
    CopyMemory c(0), MsgBkColor, 4
    Page.Paint 0, MsgX + MsgW / 2 * (1 - bpro), MsgY, MsgW * bpro + 1, MsgH + 1, argb(c(3) * bpro, c(2), c(1), c(0))
    
    DrawF = MsgBoundRect
    m = CheckMouse2
    If m = mMouseOut And MsgBtnFocus = -1 Then
        MsgBtnFocusTime = GetTickCount: MsgBtnFocus = 0
        MsgBoundRect.X = -404: MsgTipTime = 0
    End If
    If MsgBtnFocus = 1 And bpro = 0 Then
        Me.MsgButton = 1: MsgBtnFocusTime = GetTickCount: MsgBtnFocus = 0
    End If
    
    CopyMemory c(0), MsgFgColor, 4
    Page.Writes MsgContent, MsgX + 20 + MsgW / 2 * (1 - bpro), MsgY + 4, 18, argb(c(3) * bpro, c(2), c(1), c(0)), style:=FontStyleBold

    'Page.Paint 0, MsgX, MsgY, MsgW + 1, MsgH + 1, IIf(MsgTipColor = 1, argb(40, 255, 255, 255), argb(40, 0, 0, 0)), style:=1
End Sub
Public Sub SimpleMsgDrawer()
    Dim upro As Single, upro2 As Single
    upro = (GetTickCount - MsgEnterTime) / 500
    If upro > 1 Then upro = 1
    upro = Cubic(upro, 0, 1, 1, 1)
    upro2 = (GetTickCount - MsgEnterTime - 250) / 500
    If upro2 < 0 Then upro2 = 0
    If upro2 > 1 Then upro2 = 1
    upro2 = Cubic(upro2, 0, 1, 1, 1)

    Dim m As MButtonState
    
    Page.Clear 0
    PaintDC BDC, Page.CDC, alpha:=1
    
    Page.Paint 0, 0, 0, RGW + 1, RGH + 1, argb(80, 0, 0, 0)
    
    Page.Paint 2, RGW / 2, RGH - RGH / 2 * upro, MsgW * upro, MsgH * upro, argb(255 * upro, 255, 255, 255), 20, pos:=posOnCenter
    
    Page.Writes MsgTitle, RGW / 2, RGH / 2 - MsgH / 2 + 20, 18, argb(255 * upro2, 0, 0, 0), style:=FontStyleBold, align:=StringAlignmentCenter
    Page.Writes MsgContent, RGW / 2 - MsgW / 2 + 20, RGH / 2 - MsgH / 2 + 60, 18, argb(255 * upro2, 64, 64, 70), MsgW - 40, MsgH - 60 - 35 - 20, StringAlignmentCenter

    Dim pro As Single, bpro As Single
    If MsgBtnFocusTime <> 0 Then
        pro = (GetTickCount - MsgBtnFocusTime) / 500
        If pro > 1 Then pro = 1
        pro = Cubic(pro, 0, 1, 1, 1)
        If MsgBtnFocus = -1 Then
            bpro = 1 - pro
        Else
            bpro = pro
        End If
    End If
    Page.Paint 2, RGW / 2, RGH / 2 + MsgH / 2 - 20 - 20, MsgW * 0.9 * upro2, 40, argb(255 * upro2, 232 + (65 - 232) * bpro, 232 + (204 - 232) * bpro, 232 + (194 - 232) * bpro), 100, pos:=posOnCenter
    m = CheckMouse2
    Dim WSplit As Long, I As Integer, basey As Long, basex As Long
    basex = RGW / 2 - MsgW * 0.9 / 2
    basey = RGH / 2 + MsgH / 2 - 20 - 20 - 20
    WSplit = MsgW * 0.9 / (UBound(MsgButtons) + 1)
    For I = 0 To UBound(MsgButtons)
        If I < UBound(MsgButtons) Then Page.Paint 0, basex + WSplit * (I + 1) * IIf(I <> MsgBtnFocus, 1 - bpro, 1) - 2, basey + 4, 4, 32, argb(50 * (1 - bpro) * upro2 + 1, 0, 0, 0)
        Page.Writes MsgButtons(I), basex + WSplit * I * IIf(I = MsgBtnFocus, 1 - bpro, 1), basey + 6, 18, argb(255 * IIf(I <> MsgBtnFocus, 1 - bpro, 1) * upro2, 36 + (255 - 36) * IIf(I = MsgBtnFocus, bpro, 0), 36 + (255 - 36) * IIf(I = MsgBtnFocus, bpro, 0), 36 + (255 - 36) * IIf(I = MsgBtnFocus, bpro, 0)), WSplit * IIf(I <> MsgBtnFocus, 1 - bpro, 1) + 1 + (MsgW * 0.9 - WSplit) * IIf(I = MsgBtnFocus, bpro, 0), 22, StringAlignmentCenter, IIf(I = MsgBtnFocus, FontStyleBold, FontStyleRegular)
        If MsgBtnFocus = -1 And m > 0 Then
            If Mouse.X >= basex + WSplit * I And Mouse.X <= basex + WSplit * (I + 1) Then
                MsgBtnFocus = I: MsgBtnFocusTime = GetTickCount
            End If
        End If
    Next
    If MsgBtnFocus <> -1 And m >= mMouseDown Then
        Me.MsgButton = MsgBtnFocus
    End If
    If MsgBtnFocus <> -1 And m = 0 Then
        MsgBtnFocus = -1: MsgBtnFocusTime = GetTickCount
    End If
End Sub
Public Sub MsgDialog()
    Select Case MsgType
        Case 1: Call SimpleMsgDrawer
        Case 2: Call ScreenMsgDrawer
        Case 3: Call TipMsgDrawer
    End Select
End Sub
Public Sub DesignDialog()
    Page.Clear argb(60, 0, 0, 0)
    ChoosePosition = True
    
    Dim OX As Long, OY As Long
    Dim m As Integer
    m = CheckMouse(0, 0, RGW, RGH)
    
    If m >= MButtonState.mMouseDown And Mouse.button = 1 Then
        If RectMode = 0 Then
            With ChooseRect
                .Left = Mouse.X
                .top = Mouse.y
                .Right = Mouse.X
                .Bottom = Mouse.y
            End With
            RectMode = 1
        End If
        
        With ChooseRect
            .Right = Mouse.X
            .Bottom = Mouse.y
        End With
            
        If m = MButtonState.mMouseUp Then
            If ChooseRect.Right < ChooseRect.Left Then
                OX = ChooseRect.Left: ChooseRect.Left = ChooseRect.Right: ChooseRect.Right = OX
            End If
            If ChooseRect.Bottom < ChooseRect.top Then
                OY = ChooseRect.top: ChooseRect.top = ChooseRect.Bottom: ChooseRect.Bottom = OY
            End If
            RectMode = 0
        End If
    End If
    'Page.Paint 1, Mouse.X, Mouse.y, 20, 20, argb(128, 255, 255, 255), pos:=posOnCenter
    Dim yes As Boolean
    If UBound(ChooseLines) = 0 Then
        For I = 1 To 14
            yes = (Abs(Mouse.X - I / 14 * RGW) < 5)
            Page.PaintLine I / 14 * RGW, 0, I / 14 * RGW, RGH, argb(IIf(yes, 225, 30) + IIf(I = 7, 30, 0), 255, 255, 255), 2
            If yes And (m > 1) Then
                If Debug_umode = 2 Then
                    If ChooseRect.Left = ChooseRect.Right Then ChooseRect.Left = I / 14 * RGW
                    ChooseRect.Right = I / 14 * RGW
                End If
            End If
            yes = (Abs(Mouse.y - I / 14 * RGH) < 5)
            Page.PaintLine 0, I / 14 * RGH, RGW, I / 14 * RGH, argb(IIf(yes, 225, 30) + IIf(I = 7, 30, 0), 255, 255, 255), 2
            If yes And (m > 1) Then
                If Debug_umode = 2 Then
                    If ChooseRect.top = ChooseRect.Bottom Then ChooseRect.top = I / 14 * RGH
                    ChooseRect.Bottom = I / 14 * RGH
                End If
            End If
        Next
    End If
    
    For I = 1 To UBound(ChooseLines)
        Page.Paint 0, ChooseLines(I).R.Left, ChooseLines(I).R.top, ChooseLines(I).R.Right, ChooseLines(I).R.Bottom, argb(108, 0, 0, 255)
        If ChooseLines(I).mode = 0 Then
            Page.PaintLine ChooseLines(I).Data, 0, ChooseLines(I).Data, RGH, argb(255, 0, 0, 255), 1
        Else
            Page.PaintLine 0, ChooseLines(I).Data, RGW, ChooseLines(I).Data, argb(255, 0, 0, 255), 1
        End If
        If Debug_umode = 1 And m > 1 Then
            If ChooseLines(I).mode = 0 Then
                If ChooseRect.Left = ChooseRect.Right Then ChooseRect.Left = ChooseLines(I).Data
                ChooseRect.Right = ChooseLines(I).Data
            Else
                If ChooseRect.top = ChooseRect.Bottom Then ChooseRect.top = ChooseLines(I).Data
                ChooseRect.Bottom = ChooseLines(I).Data
            End If
        End If
    Next
    ReDim ChooseLines(0)
    
    Dim NRect As RECT
    NRect = ChooseRect
    
    If NRect.Right < NRect.Left Then
        OX = NRect.Left: NRect.Left = NRect.Right: NRect.Right = OX
    End If
    If NRect.Bottom < NRect.top Then
        OY = NRect.top: NRect.top = NRect.Bottom: NRect.Bottom = OY
    End If
    Page.Paint 0, NRect.Left, NRect.top, NRect.Right - NRect.Left, NRect.Bottom - NRect.top, argb(128, 255, 255, 255)
    Page.Writes "(" & NRect.Left & "," & NRect.top & ") - (" & NRect.Right - NRect.Left & "," & NRect.Bottom - NRect.top & ")", (NRect.Left + NRect.Right) / 2, (NRect.top + NRect.Bottom) / 2 - 9 / 0.75 - 1, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    
    If m = MButtonState.mMouseUp And Mouse.button = 2 Then
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText NRect.Left & "," & NRect.top & "," & NRect.Right - NRect.Left & "," & NRect.Bottom - NRect.top
        MsgBox "复制成功！", 64, "Emerald 界面设计器"
        Mouse.State = 0
    End If
    
    Page.Writes "(" & Mouse.X & "," & Mouse.y & ")", Mouse.X + 24, Mouse.y + 24, 18, argb(255, 255, 255, 255), style:=FontStyleBold
    
    Page.PaintLine 0, Mouse.y, RGW, Mouse.y, argb(200, 255, 0, 0), 2
    Page.PaintLine Mouse.X, 0, Mouse.X, RGH, argb(200, 255, 0, 0), 2
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
        Call ShowSuggestion
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
    On Error Resume Next

    If EnterTime = 0 Then EnterTime = GetTickCount
    Dim tpro As Single, tpro2 As Single
    tpro = (GetTickCount - EnterTime) / 1000
    If tpro > 1 Then tpro = 1
    tpro = Cubic(tpro, 0, 1, 1, 1)
    tpro2 = (GetTickCount - EnterTime - 1000) / 1000
    If tpro2 > 1 Then tpro2 = 1
    If tpro2 < 0 Then tpro2 = 0
    tpro2 = Cubic(tpro2, 0, 1, 1, 1)
    
    Page.Clear argb(255, 32, 32, 40)
        
    Dim pro As Single
    pro = (LoadedCount / PreLoadCount)
    
    Page.DrawImage "profile.png", GW / 2, GH / 2 - 90 * tpro, alpha:=1 * tpro, pos:=posOnCenter
    'Page.DrawImage "elogo.png", GW / 2 - 115, GH - 45, alpha:=1 * tpro2, pos:=posOnCenter
    Page.Writes "MADE WITH EMERALD", GW / 2, GH - 57, 16, argb(120 * tpro2, 242, 242, 242), align:=StringAlignmentCenter
    
    If pro = 1 Then GoTo skipLoading
    Dim cpro As Single, basey As Long, dpro As Single
    basey = GH / 2 + (GH * 0.2 / 2 + 17 - 40) * tpro2
    For I = 0 To 4
        cpro = ((GetTickCount - I * 350) Mod 4000) / 4000
        If cpro <= 0.3 Then
            cpro = cpro / 0.3
            cpro = Cubic(cpro, 0, 0, 0.47, 0.47)
        ElseIf cpro <= 0.6 Then
            cpro = (cpro - 0.3) / 0.3 * 0.06 + 0.47
        Else
            cpro = (cpro - 0.6) / 0.4
            cpro = Cubic(cpro, 0.53, 0.53, 1, 1)
        End If
        dpro = IIf(cpro > 0.5, (1 - (cpro - 0.5) / 0.5), (cpro / 0.5)) * tpro2
        Page.Paint 1, (GW + 48) * cpro - 24, basey + (GH - basey) / 2 - 23, 10, 10, argb(dpro * 255, 255 - 155 * dpro, 255 - 20 * dpro, 255 - 30 * dpro)
    Next
    
skipLoading:
    
    Page.Writes App.ProductName, GW / 2, GH / 2 + 100 + (GH * 0.2 / 2 - 160) * tpro, 18, argb(255 * tpro, 242, 242, 242), align:=StringAlignmentCenter, style:=FontStyleBold
    
    Page.Writes "正在启动 " & Int(pro * 100) & "%", GW / 2, GH / 2 + (GH * 0.2 / 2 + 17 - 40) * tpro2, 16, argb(200 * tpro2, 242, 242, 242), align:=StringAlignmentCenter

    If DisableLOGO Then
        ECore.NewTransform transFadeIn, 500
        DoneMark = True
        Exit Sub
    End If
    If PreLoadCount = LoadedCount And (GetTickCount - EnterTime) >= 3000 Then
        ECore.NewTransform transFadeIn, 500
        DoneMark = True
    End If
    
End Sub
Public Sub SaveDialog()
    Page.Paint 0, 0, 0, 900, 76, argb(255, 255, 214, 139)
    Page.DrawImage "data.png", 22, 23, alpha:=1
    Page.Writes "你希望该应用在以下位置保存数据吗？" & vbCrLf & ESave.path, 40, 130, 18, argb(255, 0, 0, 0)
    
    Page.Writes "*如果不允许，该应用可能永远保持初始状态。", 40, 400 - 110, 18, argb(255, 0, 0, 0), style:=FontStyleBold
    
    Dim MouseT As Integer
    MouseT = CheckMouse(600 - 140 * 2 + 20, 400 - 60, 100, 30)
    
    Page.Paint 0, 600 - 140 * 2 + 17, 400 - 60, 100, 30, IIf(MouseT <> 0, argb(255, 225, 184, 109), argb(255, 255, 214, 139))
    Page.Writes "允许(Y)", 600 - 140 * 2 + 17, 400 - 59, 18, argb(255, 0, 0, 0), 100, 30, StringAlignmentCenter
    If IsKeyPress(VK_Y) Or (MouseT > 1) Then Key = 1: Win.Accept
    
    MouseT = CheckMouse(600 - 140 * 1 + 10, 400 - 60, 100, 30)
    
    Page.Paint 0, 600 - 140 * 1 + 7, 400 - 60, 100, 30, IIf(MouseT <> 0, argb(255, 180, 180, 180), argb(255, 200, 200, 200))
    Page.Writes "拒绝(N)", 600 - 140 * 1 + 7, 400 - 59, 18, argb(255, 0, 0, 0), 100, 30, StringAlignmentCenter
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
    Page.IsSystem = True
    Dim LGW As Long, LGH As Long
    LGW = RGW: LGH = RGH
    RGW = Screen.Width / Screen.TwipsPerPixelX: RGH = Screen.Height / Screen.TwipsPerPixelY
    Page.Create Me
    RGW = LGW: RGH = LGH
    Page.Res.NewImages App.path & "\assets\debug"
    BDC = CreateCDC(RGW, RGH)
    If Dir(App.path & "\assets\sets\profile.png") <> "" Then
        Page.Res.newImage App.path & "\assets\sets\profile.png", 160, 160, "profile.png"
    Else
        Page.Res.newImage App.path & "\assets\debug\icon.png", 160, 160, "profile.png"
    End If
    Page.Res.newImage App.path & "\assets\debug\logo.png", 64, 64, "elogo.png"
    
    Page.TopPage = True
    

End Sub


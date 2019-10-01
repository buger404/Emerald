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
    End Select
End Sub
Public Sub SuggestionDialog()
    Page.Clear argb(0, 0, 0, 0)
    If IsKeyPress(VK_F12) Then
        If UBound(SGS) = 0 Then
            MsgBox "没有任何建议。", 48, "Emerald 建议中心"
        Else
            Dim ret As String, sRet As String
            ret = "共 " & UBound(SGS) & " 项建议。" & vbCrLf & vbCrLf
            For I = 1 To UBound(SGS)
                Select Case SGS(I).Deepth
                    Case 0: sRet = "(无关紧要)"
                    Case 1: sRet = "(需要留意)"
                    Case 2: sRet = "(*紧急)"
                    Case 3: sRet = "(*非常紧急*)"
                End Select
                ret = ret & sRet & "    " & SGS(I).Content & vbCrLf
            Next
            MsgBox ret, 48, "Emerald 建议中心"
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
    If DisableLOGO Then
        ECore.NewTransform 1, 1000
        DoneMark = True
        If bitmap <> 0 Then
            GdipDisposeImage bitmap
            GdipDeleteBrush LOGO
        End If
        Exit Sub
    End If
    If PreLoadCount = LoadedCount Then
        ECore.NewTransform 1, 1000
        DoneMark = True
    End If
    
    Page.Clear argb(255, 255, 255, 255)
    
    If LOGO = 0 Then
        Dim Graphics As Long
        GdipCreateBitmapFromScan0 GW, GH, ByVal 0, PixelFormat32bppARGB, ByVal 0, bitmap
        GdipGetImageGraphicsContext bitmap, Graphics
        GdipDrawImageRect Graphics, Page.Res.ResourceImageHandle(Page.Res.GetImage("logo.png")), GW / 2 - GH * 0.2 / 2, GH * 0.4, GH * 0.2, GH * 0.2
        GdipCreateTexture bitmap, WrapModeTile, LOGO
    End If
    
    Dim Points() As POINTF, path As Long, brush As Long, y As Long
    Dim pro As Single
    ReDim Points(1)
    
    pro = (LoadedCount / PreLoadCount)
    
    GdipCreatePath FillModeWinding, path
    GdipCreateSolidFill argb(255, 27, 27, 27), brush
    
    y = GH / 2
    
    Points(0).X = 0: Points(0).y = 0
    Points(1).X = 0: Points(1).y = y
    
    Dim Step As Long
    Step = Int(GW / 10)
    
    For X = 0 To (GW + Step) Step Step
        ReDim Preserve Points(UBound(Points) + 1)
        With Points(UBound(Points))
            .X = X
            .y = y + Sin(GetTickCount / 200 + X / 10) * (GH * 0.2 / 12)
        End With
    Next
    
    ReDim Preserve Points(UBound(Points) + 1)
    Points(UBound(Points) - 1).X = GW: Points(UBound(Points) - 1).y = y
    Points(UBound(Points)).X = GW: Points(UBound(Points)).y = 0
    
    GdipAddPathCurve path, Points(0), UBound(Points) + 1
    GdipFillPath Page.GG, brush, path
    GdipFillRectangle Page.GG, LOGO, GW / 2 - GH * 0.2 / 2, GH * 0.4, GH * 0.2, GH * 0.2
    Page.Writes Int(pro * 100) & "%", GW / 2, GH / 2 + GH * 0.2 / 2 + 20, 20, argb(255, 27, 27, 27), align:=StringAlignmentCenter
    
    GdipDeleteBrush brush
    GdipDeletePath path
    
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
    Page.TopPage = True
    
    BDC = CreateCDC(GW, GH)
End Sub


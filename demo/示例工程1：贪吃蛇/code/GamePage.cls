VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GamePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Private Type MapData
        Motion As Integer
        style As Integer
        NoClear As Boolean
    End Type
    Private Type EffectData
        time As Long
        X As Long
        y As Long
        R As Long
        g As Long
        B As Long
    End Type
    Dim PX As Long, PY As Long
    Dim Map(25, 15) As MapData, eff() As EffectData
    Dim MotionTime As Long, EnterTime As Long, FoodTime As Long
    Public DeathTime As Long
    Dim NextMotion As Integer, Speed As Long
    Dim Score As Double, DisScore As Double, HiScore As Double
'======================================
Sub CreateEffect(nX As Long, nY As Long, nR As Long, nG As Long, nB As Long)
    ReDim Preserve eff(UBound(eff) + 1)
    With eff(UBound(eff))
        .X = nX
        .y = nY
        .R = nR
        .g = nG
        .B = nB
        .time = GetTickCount
    End With
End Sub
Public Sub Update()
    '游戏每一帧的过程（绘制过程请放在此处）
    
    Page.Clear '清空画布

    '================================================================
    '   示例：绘制文字
    '   Page.Writes "我是文字", 20,20, Color:=argb(255, 0, 0, 0)
    '================================================================
    
    Page.DrawImage "background.png", Mouse.X / GW * 8 - 8, Mouse.y / GH * 8 - 8
    
    Dim pro As Single
    
    Dim DX As Long, DY As Long, X As Long, y As Long, cx As Long, cy As Long, w As Long, h As Long, Color As Long
    Dim tpro As Single, Ra As Long, xpro As Single, I As Integer
    
    DX = GW / 2 - ((UBound(Map, 1) - 1) * 30) / 2
    DY = GH / 2 - ((UBound(Map, 2) - 1) * 30) / 2
    
    xpro = (GetTickCount - MotionTime) / Speed
    If xpro > 1 Then xpro = 1
    If xpro < 0 Then xpro = 0
            
    For X = 1 To UBound(Map, 1) - 1
        For y = 1 To UBound(Map, 2) - 1
            pro = (GetTickCount - EnterTime - (X / UBound(Map, 1) * y / UBound(Map, 2)) * 1500) / 300
            If pro < 0 Then pro = 0
            If pro > 1 Then
                pro = 1
            End If
            tpro = (GetTickCount - EnterTime - (X / UBound(Map, 1) * y / UBound(Map, 2)) * 1500 - 1500) / 300
            If tpro > 1 Then tpro = 1
            If tpro < 0 Then tpro = 0
            
            pro = Cubic(pro, 0, 1, 1, 1)
            tpro = Cubic(tpro, 0, 1, 1, 1)
            cx = DX + X * 30 - 15: cy = DY + y * 30 - 15
            w = pro * 31: Ra = 31 - w
            Select Case Map(X, y).style
                Case 0
                    If tpro = 0 Then
                        Color = argb(pro * 100, pro * 255, pro * 255, pro * 255)
                    Else
                        w = 31 - tpro * 31: Ra = 31 - w
                        Color = argb(100 - tpro * 100, 255, 255, 255)
                    End If
                Case 1
                    Ra = tpro * 31
                    w = (31 - IIf(PX <> X Or PY <> y, 10, 0))
                    Color = argb((pro * 200 - IIf(PX <> X Or PY <> y, 60, 0) * pro) * tpro, 255 - pro * 255, 255 - pro * 255, 255 - pro * 255)
                    If (DeathTime > 0) And (GetTickCount Mod 500 < 250) Then Color = argb(200, 254, 84, 57)
                    Select Case Map(X, y).Motion
                        Case 1: cx = cx + xpro * 30
                        Case 2: cx = cx - xpro * 30
                        Case 3: cy = cy + xpro * 30
                        Case 4: cy = cy - xpro * 30
                    End Select
                Case 2
                    If tpro = 0 Then
                        Color = argb(pro * 200, pro * 255, pro * 255, pro * 255)
                    Else
                        w = 31 - tpro * 10: Ra = tpro * 31
                        Color = argb(200, 255, 255, 255)
                    End If
            End Select
            Page.Paint 2, cx, cy, w, w, Color, Radius:=Ra, pos:=posOnCenter
        Next
    Next
    
    Dim PPX As Single, PPY As Single
    PPX = PX: PPY = PY
    Select Case Map(PX, PY).Motion
        Case 1: PPX = PX + xpro
        Case 2: PPX = PX - xpro
        Case 3: PPY = PY + xpro
        Case 4: PPY = PY - xpro
    End Select
                    
    For X = 0 To UBound(Map, 1)
        For y = 0 To UBound(Map, 2)
            If X = 0 Or y = 0 Or X = UBound(Map, 1) Or y = UBound(Map, 2) Then
                cx = DX + X * 30 - 15: cy = DY + y * 30 - 15
                pro = (Abs(PPX - X) / UBound(Map, 1)) * 0.5 + (Abs(PPY - y) / UBound(Map, 2)) * 0.5
                pro = Cubic(pro, 0, 1, 1, 1)
                Page.Paint 0, cx, cy, 31, 31, argb(180 - 180 * pro, 255, 255, 255), pos:=posOnCenter
            End If
        Next
    Next
    
    If IsKeyPress(VK_LEFT) Then NextMotion = 2
    If IsKeyPress(VK_RIGHT) Then NextMotion = 1
    If IsKeyPress(VK_UP) Then NextMotion = 4
    If IsKeyPress(VK_DOWN) Then NextMotion = 3
    
    DisScore = DisScore + (Score - DisScore) / 10
    
    Page.Writes "得分 " & Round(DisScore) & " | 历史最高 " & HiScore, GW / 2, GH - 38, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    
    If tpro < 1 Or DeathTime > 0 Then MotionTime = GetTickCount
    If (DeathTime > 0) And (GetTickCount Mod 500 < 250) Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(20, 255, 0, 0)
    If DeathTime > 0 Then
        If GetTickCount - DeathTime >= 2000 Then
            ECore.SimpleMsg "您的得分是 " & Round(Score), "游戏结束", StrArray("我知道了")
            DeathTime = 0: ECore.NewTransform transDarkTo, , "GamePage"
        End If
    End If
    
    For I = 1 To UBound(eff)
        If I > UBound(eff) Then Exit Sub
        If I = 0 Then pro = 0: GoTo last
        pro = (GetTickCount - eff(I).time) / 1000
        If pro < 0 Then pro = 0
        If pro > 1 Then pro = 1: GoTo last
        pro = Cubic(pro, 0, 1, 1, 1)
        cx = DX + eff(I).X * 30 - 15: cy = DY + eff(I).y * 30 - 15
        Page.Paint 1, cx, cy, pro * 100, pro * 100, argb(180 - 180 * pro, eff(I).R, eff(I).g, eff(I).B), size:=2, style:=1, pos:=posOnCenter
last:
        If pro = 1 Then eff(I) = eff(UBound(eff)): ReDim Preserve eff(UBound(eff) - 1): I = I - 1
    Next
    
    If GetTickCount - MotionTime >= Speed Then
        Dim CMap() As MapData, NextC As MapData
        ReDim CMap(UBound(Map, 1), UBound(Map, 2))
        Dim fx As Long, FY As Long, NextF As MapData, ReF As Boolean
        CMap = Map
        For X = 1 To UBound(Map, 1) - 1
            For y = 1 To UBound(Map, 2) - 1
                If Map(X, y).Motion > 0 Then
                    Select Case Map(X, y).Motion
                        Case 1: NextC = Map(X + 1, y): CMap(X + 1, y) = Map(X, y)
                        Case 2: NextC = Map(X - 1, y): CMap(X - 1, y) = Map(X, y)
                        Case 3: NextC = Map(X, y + 1): CMap(X, y + 1) = Map(X, y)
                        Case 4: NextC = Map(X, y - 1): CMap(X, y - 1) = Map(X, y)
                    End Select
                    If X = PX And y = PY Then
                        Select Case Map(X, y).Motion
                            Case 1: PX = PX + 1
                            Case 2: PX = PX - 1
                            Case 3: PY = PY + 1
                            Case 4: PY = PY - 1
                        End Select
                        If PX = 0 Or PX = UBound(Map, 1) Or PY = 0 Or PY = UBound(Map, 2) Or NextC.style = 1 Then
                            BGM.Stops
                            DeathTime = GetTickCount
                            Exit Sub
                        End If
                        CMap(PX, PY).Motion = NextMotion
                        If NextC.style = 2 Then
                            Speed = Speed * 0.95
                            Score = Score + Int(Speed / 500 * 100 * (1 + Rnd * 0.5))
                            Call CreateEffect(PX, PY, 0, 0, 0)
                            ReF = True
                            GoTo NextCC
                        End If
                    End If
                    If NextC.style = 1 Then
                       Select Case Map(X, y).Motion
                            Case 1: CMap(X + 1, y).Motion = NextC.Motion: CMap(X + 1, y).NoClear = True
                            Case 2: CMap(X - 1, y).Motion = NextC.Motion: CMap(X - 1, y).NoClear = True
                            Case 3: CMap(X, y + 1).Motion = NextC.Motion: CMap(X, y + 1).NoClear = True
                            Case 4: CMap(X, y - 1).Motion = NextC.Motion: CMap(X, y - 1).NoClear = True
                       End Select
                    End If
                    If CMap(X, y).NoClear Then GoTo NextCC
                    If CMap(X, y).style = 1 Then fx = X: FY = y: NextF = CMap(X, y)
                    CMap(X, y).Motion = 0
                    CMap(X, y).style = 0
                End If
NextCC:
            Next
        Next
        If ReF Then CMap(fx, FY) = NextF
        For X = 1 To UBound(Map, 1) - 1
            For y = 1 To UBound(Map, 2) - 1
                CMap(X, y).NoClear = False
                Map(X, y) = CMap(X, y)
            Next
        Next
        MotionTime = GetTickCount
    End If
    
    If GetTickCount - FoodTime >= Speed * 10 Then
        Randomize
        X = Int(Rnd * (UBound(Map, 1) - 3)) + 3
        y = Int(Rnd * (UBound(Map, 2) - 3)) + 3
        If Map(X, y).style <> 1 Then
            Map(X, y).style = 2
            Call CreateEffect(X, y, 255, 255, 255)
            FoodTime = GetTickCount
        End If
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    ReDim eff(0)
    Erase Map
    EnterTime = GetTickCount
    Map(1, 1).style = 1
    PX = 1: PY = 1
    Randomize
    Dim X As Long, y As Long
    For I = 1 To 10
        X = Int(Rnd * (UBound(Map, 1) - 3)) + 3
        y = Int(Rnd * (UBound(Map, 2) - 3)) + 3
        Map(X, y).style = 2
        Call CreateEffect(X, y, 255, 255, 255)
    Next
    NextMotion = 1
    Map(1, 1).Motion = 1
    DeathTime = 0
    If Score > HiScore Then
        ESave.PutData "Record", Score
        ESave.Save
    End If
    
    HiScore = Val(ESave.GetData("Record"))
    Speed = 500: Score = 0
    'ECore.FancyMode = True
End Sub
Public Sub Leave()
    '页面离开事件
    'ECore.FancyMode = False
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    ReDim eff(0)
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets", 0.7
    '创建页面
    ECore.Add Page, "GamePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

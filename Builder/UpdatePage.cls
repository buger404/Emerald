VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UpdatePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage
    Dim Step As Long
    Dim PIndex As Long
    Private Type CopyList
        path As String
        mark As Integer '    0=新增 1=变化 2=用户修改
        Rep As Boolean
        incrase As Long
    End Type
    Dim FileList() As CopyList
    Dim BackupSwitch As Boolean
    Dim WarnStr As String
'======================================
Public Sub CheckFiles()
    Dim f As String, fl() As String
    ReDim FileList(0), fl(0)
    
    f = Dir(App.path & "\core\")
    Do While f <> ""
        DoEvents
        ReDim Preserve fl(UBound(fl) + 1)
        fl(UBound(fl)) = f
        f = Dir()
    Loop
    
    For i = 1 To UBound(fl)
        If Dir(Cmd & "\core\" & fl(i)) = "" Then
            ReDim Preserve FileList(UBound(FileList) + 1)
            With FileList(UBound(FileList))
                .path = fl(i)
                .mark = 0
                .incrase = FileLen(App.path & "\core\" & fl(i))
                .Rep = True
            End With
        End If
    Next
    
    Dim om As Boolean, t As String
    
    For i = 1 To UBound(fl)
        om = False
        t = 0
        If Dir(Cmd & "\core\" & fl(i)) <> "" Then
            If Dir(Cmd & "\.emr\cache\" & fl(i)) <> "" Then
                Open Cmd & "\.emr\cache\" & fl(i) For Input As #1
                Line Input #1, t
                Close #1
                If Val(t) <> FileLen(Cmd & "\core\" & fl(i)) Then om = True
            End If
            If (om = False) And (FileLen(Cmd & "\core\" & fl(i)) <> FileLen(App.path & "\core\" & fl(i))) Then
                ReDim Preserve FileList(UBound(FileList) + 1)
                With FileList(UBound(FileList))
                    .path = fl(i)
                    .mark = 1
                    .incrase = FileLen(App.path & "\core\" & fl(i)) - t
                    .Rep = True
                End With
            End If
        End If
    Next
    
    For i = 1 To UBound(fl)
        om = True
        t = 0
        If Dir(Cmd & "\core\" & fl(i)) <> "" Then
            If Dir(Cmd & "\.emr\cache\" & fl(i)) <> "" Then
                Open Cmd & "\.emr\cache\" & fl(i) For Input As #1
                Line Input #1, t
                Close #1
                If Val(t) <> FileLen(Cmd & "\core\" & fl(i)) Then om = False
            End If
            If (om = False) Then
                ReDim Preserve FileList(UBound(FileList) + 1)
                With FileList(UBound(FileList))
                    .path = fl(i)
                    .mark = 2
                    .incrase = FileLen(App.path & "\core\" & fl(i)) - t
                    .Rep = False
                End With
            End If
        End If
    Next
End Sub
Public Sub Update2()
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布
    
    Page.Writes "替换代码文件", 50, 100, 24, argb(255, 27, 27, 27)
    Page.Writes "请选择您需要替换的代码", 50, 140, 18, argb(150, 27, 27, 27)
    
    Dim Color As Long
    
    For i = PIndex + 1 To PIndex + 5
        If i > UBound(FileList) Then Exit For
        
        Page.Paint 1, 50, 170 + (i - PIndex) * 40, 24, 24, IIf(FileList(i).Rep, argb(255, 95, 229, 221), argb(255, 222, 222, 222))
        If CheckMouse2 = mMouseUp Then FileList(i).Rep = Not FileList(i).Rep
        
        Select Case FileList(i).mark
            Case 0
                Color = argb(180, 27, 27, 27)
                Page.Paint 0, GW - 150, 170 + (i - PIndex) * 40, 90, 24, argb(255, 95, 229, 221)
                Page.Writes "新增", GW - 150, 170 + (i - PIndex) * 40, 16, argb(255, 255, 255, 255), 90, 24, StringAlignmentCenter
            Case 1
                Color = argb(255, 240, 127, 34)
                Page.Paint 0, GW - 150, 170 + (i - PIndex) * 40, 90, 24, Color
                Page.Writes "有更新", GW - 150, 170 + (i - PIndex) * 40, 16, argb(255, 255, 255, 255), 90, 24, StringAlignmentCenter
            Case 2
                Color = argb(255, 232, 76, 61)
                Page.Paint 0, GW - 150, 170 + (i - PIndex) * 40, 90, 24, Color
                Page.Writes "您修改过", GW - 150, 170 + (i - PIndex) * 40, 16, argb(255, 255, 255, 255), 90, 24, StringAlignmentCenter
        End Select
        
        Page.Writes IIf(FileList(i).incrase > 0, "+", "") & FileList(i).incrase & " bytes", GW - 170, 170 + (i - PIndex) * 40, 18, argb(140, 27, 27, 27), style:=FontStyleBold, align:=StringAlignmentFar
        Page.Writes FileList(i).path, 80, 170 + (i - PIndex) * 40, 18, Color, style:=IIf(FileList(i).mark = 2, FontStyleBold, 0)
        
        If i <> PIndex + 5 Then Page.PaintLine 50, 170 + (i - PIndex) * 40 + 32, GW - 50, 170 + (i - PIndex) * 40 + 32, argb(255, 222, 222, 222), 1
    Next
    
    Page.Writes (Int(PIndex / 5) + 1) & " / " & (Int(UBound(FileList) / 5) + 1), GW / 2, GH - 140, 16, argb(150, 27, 27, 27), align:=StringAlignmentCenter
    If Page.ShowColorButton(2, GW / 2 - 80 - 20, GH - 140, 40, 22, "<", argb(150, 27, 27, 27), argb(255, 222, 222, 222), argb(130, 95, 229, 221), size:=16, Radius:=60) = 3 Then
        If PIndex > 0 Then PIndex = PIndex - 5: ECore.NewTransform
    End If
    If Page.ShowColorButton(2, GW / 2 + 80 - 20, GH - 140, 40, 22, ">", argb(150, 27, 27, 27), argb(255, 222, 222, 222), argb(130, 95, 229, 221), size:=16, Radius:=60) = 3 Then
        If (Int(PIndex / 5) + 1) < (Int(UBound(FileList) / 5) + 1) Then PIndex = PIndex + 5: ECore.NewTransform
    End If
    
    If Page.ShowColorButton(0, GW - 120 - 50, GH - 80, 120, 40, "替换", argb(150, 27, 27, 27), argb(255, 222, 222, 222), argb(130, 95, 229, 221), size:=18) = 3 Then
        If Dialog("警告", "即将替换所选文件。" & IIf(BackupSwitch, "并创建备份。", ""), "好", "等等！") <> 1 Then Exit Sub
        ECore.NewTransform transFadeIn, 700, "WaitPage"
        
        Dim Backstr As String
        If BackupSwitch Then
            Backstr = year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日" & Hour(Now) & "时" & Minute(Now) & "分" & Second(Now) & "秒"
            If Dir(Cmd & "\.emr\backup\" & Backstr, vbDirectory) = "" Then MkDir Cmd & "\.emr\backup\" & Backstr
            CopyInto Cmd & "\core", Cmd & "\.emr\backup\" & Backstr
            If Dir(Cmd & "\.emr\backup\" & Backstr & "\cache", vbDirectory) = "" Then MkDir Cmd & "\.emr\backup\" & Backstr & "\cache"
            CopyInto Cmd & "\.emr\cache", Cmd & "\.emr\backup\" & Backstr & "\cache"
            If Dir(Cmd & "\.emr\backup\" & Backstr & "\project", vbDirectory) = "" Then MkDir Cmd & "\.emr\backup\" & Backstr & "\project"
            FileCopy Cmd & "\" & Dir(Cmd & "\*.vbp"), Cmd & "\.emr\backup\" & Backstr & "\project\" & Dir(Cmd & "\*.vbp")
        End If
        
        Dim NoUpdate As Boolean, AddList As String, FileMark As String
        For i = 1 To UBound(FileList)
            If FileList(i).Rep Then
                FileCopy App.path & "\core\" & FileList(i).path, Cmd & "\core\" & FileList(i).path
                Open Cmd & "\.emr\cache\" & FileList(i).path For Output As #1
                Print #1, FileLen(Cmd & "\core\" & FileList(i).path)
                Close #1
                FileMark = ""
                If FileList(i).mark = 0 Then
                    If LCase(FileList(i).path) Like "*.bas" Then FileMark = "Module"
                    If LCase(FileList(i).path) Like "*.cls" Then FileMark = "Class"
                    If LCase(FileList(i).path) Like "*.frm" Then FileMark = "Form"
                    
                    If FileMark <> "" Then
                        If FileMark = "Form" Then
                            AddList = AddList & FileMark & "=" & "Core\" & FileList(i).path & vbCrLf
                        Else
                            AddList = AddList & FileMark & "=" & Split(FileList(i).path, ".")(0) & "; Core\" & FileList(i).path & vbCrLf
                        End If
                    End If
                End If
            Else
                NoUpdate = True
            End If
            Call FakeSleep(30)
        Next
        CopyInto App.path & "\assets\debug", Cmd & "\assets\debug"
        CopyInto App.path & "\framework", Cmd
        
        If AddList <> "" Then
            Dim tempVbp As String, tempRead As String
            Open Cmd & "\" & Dir(Cmd & "\*.vbp") For Input As #1
            Do While Not EOF(1)
                Line Input #1, tempRead
                If Left(tempRead, 7) = "Startup" Then
                    tempVbp = tempVbp & AddList & tempRead
                Else
                    tempVbp = tempVbp & tempRead & vbCrLf
                End If
            Loop
            Close #1
            Open Cmd & "\" & Dir(Cmd & "\*.vbp") For Output As #1
            Print #1, tempVbp
            Close #1
        End If
        
        Open Cmd & "\.emerald" For Output As #1
        Print #1, Version 'version
        Print #1, Now 'Update Time
        Print #1, NoUpdate
        Close #1
        
        Dialog "完成", "你的工程已经更新了！", "好"
        Unload MainWindow: End
    End If
    
    Page.Writes "帮助我选择...", 50, GH - 30 - 40, 16, argb(150, 27, 27, 27), 200, 20, style:=FontStyleUnderline
    If CheckMouse2 = mMouseUp Then
        Call CheckFiles
        ECore.NewTransform
    End If
End Sub
Public Sub Update()
    '绘制事件
    If Step = 1 Then Call Update2: Exit Sub
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '清空画布
    
    Page.Writes "欢迎使用工程更新向导", 50, 100, 24, argb(255, 27, 27, 27)
    Page.Writes "该向导将会帮助你顺利更新Emerald有关代码。", 50, 140, 18, argb(150, 27, 27, 27)
    
    Dim lead As String, leadstr As String
    lead = format(Version - Val(AppInfo(0)), "00000000")
    If Left(lead, 2) <> "00" Then leadstr = leadstr & Val(Left(lead, 2)) & "年"
    If Right(Left(lead, 4), 2) <> "00" Then leadstr = leadstr & Val(Right(Left(lead, 4), 2)) & "个月"
    If Right(Left(lead, 6), 2) <> "00" Then leadstr = leadstr & Val(Right(Left(lead, 6), 2)) & "天"
    If leadstr = "" Then
        If Right(Left(lead, 8), 2) <> "00" Then leadstr = leadstr & Val(Right(Left(lead, 8), 2)) & "项更新"
    Else
        leadstr = leadstr & "，若干项更新"
    End If
    
    Page.Writes "最新版本：  " & Version & "（领先" & leadstr & "）", 50, 220, 18, argb(220, 27, 27, 27)
    Page.Writes "工程版本：  " & Val(AppInfo(0)), 50, 250, 18, argb(130, 27, 27, 27)
    
    If WarnStr <> "" Then
        Page.Paint 0, 0, 295, GW, 130, argb(255, 255, 255, 255)
        Page.Paint 0, 0, 295, 10, 130, argb(255, 255, 205, 66)
        Page.Writes WarnStr, 50, 315, 16, argb(130, 27, 27, 27)
    End If
    
    If Page.ShowColorButton(0, GW - 150 - 50, GH - 100, 150, 50, "立即更新", argb(150, 27, 27, 27), argb(255, 222, 222, 222), argb(130, 95, 229, 221), size:=18) = 3 Then
        Call CheckFiles
        ECore.NewTransform , 700
        PIndex = 0
        Step = 1
    End If
    
    Page.Paint 1, 50, GH - 30 - 40 - 40, 24, 24, IIf(BackupSwitch, argb(255, 95, 229, 221), argb(255, 222, 222, 222))
    If CheckMouse2 = mMouseUp Then BackupSwitch = Not BackupSwitch
    Page.Writes "创建备份", 50 + 35, GH - 30 - 40 - 40 + 1, 16, argb(150, 27, 27, 27)
    
    Page.Writes "我希望了解更新内容", 50, GH - 30 - 40, 16, argb(150, 27, 27, 27), 200, 20, style:=FontStyleUnderline
    If CheckMouse2 = mMouseUp Then ShellExecuteA 0, "open", "https://github.com/Red-Error404/Emerald/commits/master", "", "", SW_SHOW
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub Enter()
    '页面进入事件
    Step = 0: BackupSwitch = True
End Sub
Public Sub GetWarnStr()
    WarnStr = ""
    If Val(AppInfo(0)) < 19051004 Then
        WarnStr = WarnStr & "资源加载函数已经迁移(Page->Page.Res)" & vbCrLf
    End If
    If Val(AppInfo(0)) < 19051110 Then
        WarnStr = WarnStr & "窗口鼠标检测出现问题（请参照DebugSwitch模块里的注释修改代码）" & vbCrLf
        WarnStr = WarnStr & "画布清空机制修改（请在你的绘图过程加上Page.Clear）" & vbCrLf
    End If
    If Val(AppInfo(0)) < 19060706 Then
        WarnStr = WarnStr & "存档机制修改（详情查看DebugSwitch模块中的注释）" & vbCrLf
        WarnStr = WarnStr & "动画的加入需要你修改部分代码（详情查看DebugSwitch模块中的注释）" & vbCrLf
    End If
    If Val(AppInfo(0)) < 19062303 Then
        WarnStr = WarnStr & "文本框控件的引入需要你添加一些代码（详情查看DebugSwitch模块中的注释）" & vbCrLf
    End If
End Sub
Public Sub Leave()
    '页面离开事件
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
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.path & "\assets\debug"
    '创建页面
    ECore.Add Page, "UpdatePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WelcomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Public Page As GPage
    Public PageIndex As Integer, Cliped As Integer
    Public CMError As String
    Public StartTime As Long
    Dim ReadProject As Boolean
    Private Type UpdateFiles
        Name As String
        IncraseByte As String
        Kind As Integer
        Image As String
        ReplaceMark As Boolean
    End Type
    Private Type BackupItem
        Name As String
        UseTime As Long
        Date As String
        FileCount As Long
        size As Long
    End Type
    Dim BackupList() As BackupItem, BackupIndex As Integer
    Dim UFiles() As UpdateFiles, FileIndex As Integer
    Dim BackupSwitch As Boolean, ReplaceAccept As Long
    Dim PackState As Long
    Public PackText As String
'======================================
Public Sub ScanBackup()
    Dim File As String, Backup As EmrBackup
    File = Dir(OPath & "\.emr\*.emrbackup")
    ReDim BackupList(0)
    Do While File <> ""
        Open OPath & "\.emr\" & File For Binary As #1
        Get #1, , Backup
        Close #1
        If Backup.AFileHeader = "Emerald Project Backup" Then
            ReDim Preserve BackupList(UBound(BackupList) + 1)
            With BackupList(UBound(BackupList))
                .Date = Backup.Date
                .FileCount = UBound(Backup.Files)
                .Name = Split(File, ".emrbackup")(0)
                If .Name = "" Then .Name = "null"
                .size = FileLen(OPath & "\.emr\" & File)
            End With
        End If
        File = Dir()
    Loop
    
    If Dir(OPath & "\.emr\backup", vbDirectory) <> "" Then
        ReDim Preserve BackupList(UBound(BackupList) + 1)
        With BackupList(UBound(BackupList))
            .Date = "未知"
            .FileCount = -1
            .Name = "远古版本的Emerald创建的备份"
            .size = -1
        End With
    End If
End Sub
Public Sub WriteCache()
    Dim Cache As EmrBackup
    Dim File As String, Check As String, MD5 As New MD5
    ReDim Cache.Files(0)
    
    Cache.AFileHeader = "Emerald Project Cache"
    File = Dir(OPath & "\Core\")
    Do While File <> ""
        ReDim Preserve Cache.Files(UBound(Cache.Files) + 1)
        Check = MD5.Md5_File_Calc(OPath & "\Core\" & File)
        With Cache.Files(UBound(Cache.Files))
            .MD5Check = Check
            .path = "Core\" & File
        End With
        File = Dir()
        ECore.Display: DoEvents
    Loop
    
    Open OPath & "\.emr\.cache" For Binary As #1
    Put #1, , Cache
    Close #1
    
    Set MD5 = Nothing
End Sub
Public Sub ReplaceProject()
    Dim FUpdate As Boolean, Backup As EmrBackup
    Dim File As String, Data() As Byte, Check As String, MD5 As New MD5
    ReDim Backup.Files(0)
    
    Backup.AFileHeader = "Emerald Project Backup"
    Backup.Date = Now

    File = Dir(OPath & "\Core\")
    Do While File <> ""
        ReDim Preserve Backup.Files(UBound(Backup.Files) + 1)
        Check = MD5.Md5_File_Calc(OPath & "\Core\" & File)
        With Backup.Files(UBound(Backup.Files))
            .MD5Check = Check
            .path = "Core\" & File
            ReDim Data(FileLen(OPath & "\Core\" & File) - 1)
            Open OPath & "\Core\" & File For Binary As #1
            Get #1, , Data
            Close #1
            .Data = Data
        End With
        File = Dir()
        ECore.Display: DoEvents
    Loop
    ReDim Preserve Backup.Files(UBound(Backup.Files) + 1)
    Check = ""
    With Backup.Files(UBound(Backup.Files))
        .MD5Check = Check
        .path = Dir(OPath & "\*.vbp")
        ReDim Data(FileLen(OPath & "\" & Dir(OPath & "\*.vbp")) - 1)
        Open OPath & "\" & Dir(OPath & "\*.vbp") For Binary As #1
        Get #1, , Data
        Close #1
        .Data = Data
    End With
        
    For I = 1 To UBound(UFiles)
        If UFiles(I).ReplaceMark Then
            FileCopy App.path & "\Core\" & UFiles(I).Name, OPath & "\Core\" & UFiles(I).Name
        Else
            FUpdate = True
        End If
        ECore.Display: DoEvents
    Next
    
    Call WriteCache
    If BackupSwitch Then
        Open VBA.Environ("temp") & "\emrbackup" For Binary As #1
        Put #1, , Backup
        Close #1
    End If
    
    For I = 1 To 100
        ECore.Display: DoEvents
    Next
    
    Dim BranchId As String
    BranchId = MD5.Md5_File_Calc(VBA.Environ("temp") & "\emrbackup")
    
    EmrPC.FUpdate = FUpdate
    EmrPC.Version = Version
    If BackupSwitch Then EmrPC.LastBranch = BranchId
    
    If BackupSwitch Then
        Open OPath & "\.emr\" & BranchId & ".emrbackup" For Binary As #1
        Put #1, , Backup
        Close #1
    End If
    
    Open OPath & "\.emerald" For Binary As #1
    Put #1, , EmrPC
    Close #1
    
    ReDim UFiles(0)
    Call UpdateProject
    
    CopyInto App.path & "\assets\debug\", OPath & "\assets\debug"
    
    Set MD5 = Nothing
End Sub
Public Sub BuildInstaller()
    Dim Name As String, Describe As String, GVersion As String
    Dim tempr As String, targetEXE As String
    
    targetEXE = App.path & "\Builder.exe"
    
    WelcomePage.PackText = "正在收集工程信息 ..."
    Open OPath & "\" & Dir(OPath & "\*.vbp") For Input As #1
    Do While Not EOF(1)
        Line Input #1, tempr
        If InStr(tempr, "VersionProductName") = 1 Then Name = Split(tempr, """")(1)
        If InStr(tempr, "VersionFileDescription") = 1 Then Describe = Split(tempr, """")(1)
        If InStr(tempr, "VersionCompanyName") = 1 Then Maker = Split(tempr, """")(1)
        If InStr(tempr, "MajorVer") = 1 Then GVersion = GVersion & Split(tempr, "=")(1) & "."
        If InStr(tempr, "MinorVer") = 1 Then GVersion = GVersion & Split(tempr, "=")(1) & "."
        If InStr(tempr, "RevisionVer") = 1 Then GVersion = GVersion & Split(tempr, "=")(1)
    Loop
    Close #1
    MakePackage OPath, EmrPC.Maker, EmrPC.Name, GVersion, Describe, 0
    CreateFolder GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\"
    WelcomePage.PackText = "正在建立安装包数据库 ..."
    If Dir(GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & EmrPC.Name & " - Installer.exe") <> "" Then Kill GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & EmrPC.Name & " - Installer.exe"
    Open VBA.Environ("temp") & "\copyemr.cmd" For Output As #1
    Print #1, "@echo off"
    Print #1, "echo Emerald 安装包打包工具 , 版本号: " & Version
    Print #1, "echo 正在创建安装程序..."
    Print #1, "ping localhost -n 3 > nul"
    Print #1, "copy """ & targetEXE & """ /b + """ & VBA.Environ("temp") & "\emrpack.empack"" /b """ & GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & EmrPC.Name & " - Installer.exe"""
    Close #1
    ShellExecuteA 0, "open", VBA.Environ("temp") & "\copyemr.cmd", "", "", SW_SHOW
    Do While Dir(GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & EmrPC.Name & " - Installer.exe") = ""
        Sleep 10: DoEvents
        ECore.Display
    Loop
    ShellExecuteA 0, "open", "explorer.exe", "/select,""" & GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & EmrPC.Name & " - Installer.exe" & """", "", SW_SHOW
End Sub
Public Sub UpdateProject()
    If EmrPC.Version >= Version And Not EmrPC.FUpdate Then Exit Sub
    
    Dim File As String, Kind As Integer
    Dim Cache As EmrBackup
    Dim Check As String, Check2 As String, MD5 As New MD5, size As Long
    If Dir(OPath & "\.emr\.cache") <> "" Then
        Open OPath & "\.emr\.cache" For Binary As #1
        Get #1, , Cache
        Close #1
    End If
    File = Dir(App.path & "\Core\")
    Do While File <> ""
        If File = "Core.bas" Then GoTo NextFile
        '1=新增，2=更新，3=被修改
        Kind = IIf(IsExitAFile(OPath & "\Core\" & File), 2, 1)
        If Kind = 2 Then
            Check = MD5.Md5_File_Calc(OPath & "\Core\" & File)
            Check2 = MD5.Md5_File_Calc(App.path & "\Core\" & File)
            If Cache.AFileHeader = "Emerald Project Cache" Then
                For I = 1 To UBound(Cache.Files)
                    If Cache.Files(I).path = "Core\" & File Then
                        Kind = IIf(Cache.Files(I).MD5Check = Check, 2, 3)
                        On Error Resume Next
                        size = IIf(Kind = 3, -1, FileLen(App.path & "\Core\" & File) - FileLen(OPath & "\Core\" & File))
                        Exit For
                    End If
                Next
            Else
                size = FileLen(App.path & "\Core\" & File) - FileLen(OPath & "\Core\" & File)
            End If
            If Check = Check2 Then Kind = 0
        End If
        If Kind = 1 Then size = FileLen(App.path & "\Core\" & File)
        If Kind <> 0 Then
            ReDim Preserve UFiles(UBound(UFiles) + 1)
            With UFiles(UBound(UFiles))
                .IncraseByte = size
                .Kind = Kind
                .ReplaceMark = IIf(Kind = 3, False, True)
                .Name = File
                If .Name = "AboutMe.bas" Then .Image = "info.png"
                If .Name = "EmeraldWindow.frm" Then .Image = "windows.png"
                If .Name = "AeroEffect.bas" Then .Image = "windows.png"
                If .Name = "Animations.bas" Then .Image = "fireworks.png"
                If .Name = "Bass.bas" Then .Image = "headphones.png"
                If .Name = "BMEA_Engine.bas" Then .Image = "security.png"
                If .Name = "Core.bas" Then .Image = "cpu.png"
                If .Name = "GCore.bas" Then .Image = "cpu.png"
                If .Name = "Debuginfo.frm" Then .Image = "debugging.png"
                If .Name = "GDebug.cls" Then .Image = "debugging.png"
                If .Name = "DebugSwitch.bas" Then .Image = "info.png"
                If .Name = "DebugWindow.frm" Then .Image = "debugging.png"
                If .Name = "GAnimation.cls" Then .Image = "fireworks.png"
                If .Name = "GCrashBox.cls" Then .Image = "box.png"
                If .Name = "Gdiplus.bas" Then .Image = "paint-brush.png"
                If .Name = "GFont.cls" Then .Image = "paint-brush.png"
                If .Name = "GPage.cls" Then .Image = "paint-brush.png"
                If .Name = "GResource.cls" Then .Image = "paint-brush.png"
                If .Name = "GMan.cls" Then .Image = "cpu.png"
                If .Name = "GMusic.cls" Then .Image = "headphones.png"
                If .Name = "GMusicList.cls" Then .Image = "headphones.png"
                If .Name = "GSaving.cls" Then .Image = "folder.png"
                If .Name = "GSysPage.cls" Then .Image = "cpu.png"
                If .Name = "ResPool.bas" Then .Image = "cpu.png"
            End With
        End If
NextFile:
        File = Dir()
    Loop
    
    Set MD5 = Nothing
End Sub
Public Sub DrawAbout()
    If GetTickCount - StartTime >= 8000 And StartTime <> 0 Then
        StartTime = 0
        Call CheckOnLineUpdate
        ECore.NewTransform transFadeIn, 1000
    End If

    Page.Writes "版本", 381, 135, 18, argb(255, 35, 170, 242)
    Page.Writes "当前Builder版本号：" & Version, 381, 170, 18, argb(255, 242, 242, 242)
    Page.Writes "当前Emerald版本号：" & Version, 381, 200, 18, argb(255, 242, 242, 242)
    Page.Writes "关于", 381, 265, 18, argb(255, 35, 170, 242)
    Page.Writes "由Error404(QQ 1361778219)制作", 381, 305, 18, argb(255, 242, 242, 242)
    Page.Writes "支持", 381, 365, 18, argb(255, 35, 170, 242)
    Page.Writes "QQ群 : Inter.Net / Redstone", 381, 405, 18, argb(255, 242, 242, 242)
    Page.Writes "南笙", 381, 435, 18, argb(255, 242, 242, 242)
    Page.Writes "更新计划", 381, 505, 18, argb(255, 35, 170, 242)
    If NewVersion = 0 Then
        Page.ShowLoading 386, 545, 20, 20, 4, argb(255, 35, 170, 242), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
        Page.Writes "正在连接服务器...", 416, 547, 18, argb(255, 242, 242, 242)
    ElseIf NewVersion = 3 Then
        Page.Writes "连接超时。", 381, 547, 18, argb(255, 254, 84, 57)
    ElseIf NewVersion > Version Then
        Page.Writes "更新可用，点击此处下载。", 381, 547, 18, argb(255, 35, 170, 242), 200, 20, style:=FontStyleUnderline
        If CheckMouse2 = mMouseUp Then ShellExecuteA 0, "open", "https://github.com/buger404/Emerald/releases", "", "", SW_SHOW
    Else
        Page.Writes "已经是最新版了。", 381, 547, 18, argb(255, 242, 242, 242)
    End If
    
    Page.Writes "使用帮助", 381, 605, 18, argb(255, 35, 170, 242)
    m = CheckMouse(381, 645, 170, 20)
    Page.Writes "Github : Emerald Wiki", 381, 645, 18, IIf(m = 0, argb(255, 242, 242, 242), argb(255, 35, 170, 242)), 200, 20, style:=FontStyleUnderline
    If m = 3 Then ShellExecuteA 0, "open", "https://github.com/buger404/Emerald/wiki", "", "", SW_SHOW
    m = IIf(m = 0, MousePointerConstants.vbDefault, MousePointerConstants.vbArrowQuestion)
    If Screen.MousePointer <> m Then Screen.MousePointer = m
End Sub
Public Sub DrawProject()
    Page.Writes "游戏名称", 380, 135, 18, argb(255, 35, 170, 242)
    Page.Writes EmrPC.Name, 381, 170, 18, argb(255, 242, 242, 242)
    Page.Writes "制作人", 380, 235, 18, argb(255, 35, 170, 242)
    'Page.Writes EmrPC.Maker, 381, 270, 18, argb(255, 242, 242, 242)
    If Page.ShowEdit(EmrPC.Maker, 0, 380, 270 - 5 + 2, 280, 30, argb(255, 242, 242, 242), argb(0, 255, 255, 255), argb(50, 242, 242, 242), argb(255, 35, 170, 242), size:=18) = 4 Then
        Open OPath & "\.emerald" For Binary As #1
        Put #1, , EmrPC
        Close #1
    End If
    
    Dim m As Long
    
    Page.Writes "使用的Emerald版本号" & IIf(UBound(UFiles) > 0, " (更新可用)", ""), 380, 335, 18, IIf(UBound(UFiles) > 0, argb(255, 254, 84, 57), argb(255, 35, 170, 242))
    Page.Writes "V" & EmrPC.Version & IIf(UBound(UFiles) > 0, " -> V" & Version, ""), 381, 370, 18, argb(255, 242, 242, 242)
    
    If UBound(UFiles) = 0 Then Exit Sub
    If ReplaceAccept = -1 Then
        Page.ShowLoading 600, 460, 128, 128, 3, argb(255, 35, 170, 242), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
        Exit Sub
    End If
    
    m = CheckMouse(381, 400, 570, 20)
    Page.Writes "最后一次更新【吧", 381, 400, 18, IIf(m > 0, argb(255, 35, 170, 242), argb(255, 242, 242, 242)), 570, 20, style:=FontStyleUnderline
    If m = 3 Then ShellExecuteA 0, "open", "https://github.com/buger404/Emerald/releases", "", "", SW_SHOW
    m = IIf(m = 0, MousePointerConstants.vbDefault, MousePointerConstants.vbArrowQuestion)
    If Screen.MousePointer <> m Then Screen.MousePointer = m
    
    Dim UColor(3) As Long, UText(3) As String
    UColor(1) = argb(255, 67, 219, 210): UText(1) = "新增"
    UColor(2) = argb(255, 35, 170, 242): UText(2) = "更新"
    UColor(3) = argb(255, 254, 84, 57): UText(3) = "替换"
    
    For I = FileIndex * 6 + 1 To FileIndex * 6 + 5 + 1
        If I > UBound(UFiles) Or I < 1 Then Exit For
        Page.Paint 0, 372, 430 + 30 * (I - FileIndex * 6) - 5 - 8, 558, 31, IIf(I Mod 2 = 0, argb(0, 242, 242, 242), argb(50, 255, 255, 255))
        Page.Paint 1, 372 + 5, 431 + 30 * (I - FileIndex * 6) + 1 - 8, 15, 15, IIf(UFiles(I).ReplaceMark, argb(255, 254, 84, 57), argb(255, 168, 168, 168))
        If CheckMouse2 = mMouseUp Then UFiles(I).ReplaceMark = Not UFiles(I).ReplaceMark
        Page.Writes UFiles(I).Name, 501 + 15, 428 + 30 * (I - FileIndex * 6) - 8, 18, argb(255, 255, 255, 255)
        Page.Paint 0, 382 + 15, 430 + 30 * (I - FileIndex * 6) - 8, 80, 20, UColor(UFiles(I).Kind)
        Page.DrawImageEx UFiles(I).Image, 388 + 80 + 5 + 15, 430 + 30 * (I - FileIndex * 6) - 8, 20, 20
        Page.Writes UText(UFiles(I).Kind), 382 + 15, 428 + 30 * (I - FileIndex * 6) - 8, 16, argb(255, 255, 255, 255), 80, 20, StringAlignmentCenter, FontStyleBold
        Page.Writes IIf(UFiles(I).IncraseByte > 0, "+", "") & UFiles(I).IncraseByte & " 字节", GW - 80, 428 + 30 * (I - FileIndex * 6) - 8, 18, argb(255, 192, 192, 192), align:=StringAlignmentFar
    Next
    
    Dim max As Long
    max = Int(UBound(UFiles) / 6) + 1
    Page.Writes FileIndex + 1 & "/" & max, 600, 430 + 30 * 7 + 5, 18, argb(255, 242, 242, 242), 120, 20, align:=StringAlignmentCenter
    If FileIndex + 1 > 1 Then
        If Page.ShowColorButton(1, 600, 430 + 30 * 7 + 5, 20, 20, "<", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=16) = 3 Then
            FileIndex = FileIndex - 1
            ECore.NewTransform
        End If
    End If
    If FileIndex + 1 < max Then
        If Page.ShowColorButton(1, 700, 430 + 30 * 7 + 5, 20, 20, ">", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=16) = 3 Then
            FileIndex = FileIndex + 1
            ECore.NewTransform
        End If
    End If
    If FileIndex + 1 > max Then FileIndex = max - 1
    
    Page.ShowColorCheckBox BackupSwitch, 384, GH - 80, 200, 20, "同时创建备份", argb(255, 242, 242, 242), argb(255, 35, 170, 242), argb(255, 242, 242, 242), 18

    If Page.ShowColorButton(0, GW - 180, GH - 82, 120, 30, IIf(ReplaceAccept = 0, "替换代码", "确认替换 (" & Round(5 - (GetTickCount - ReplaceAccept) / 1000) & "s)"), argb(255, 255, 255, 255), IIf(ReplaceAccept = 0, argb(255, 35, 170, 242), argb(255, 254, 84, 57)), IIf(ReplaceAccept = 0, argb(255, 75, 210, 255), argb(255, 255, 124, 97)), 40, size:=18) = 3 Then
        If ReplaceAccept > 0 Then
            ReplaceAccept = -1
            Call ReplaceProject
            Call ScanBackup
            ReplaceAccept = 0
            MsgBox "Emerald成功替换了旧版本代码。" & IIf(BackupSwitch, vbCrLf & "并成功创建了旧版本的备份 : " & OPath & "\.emr\" & EmrPC.LastBranch & ".emrbackup", ""), 64
        Else
            ReplaceAccept = GetTickCount
        End If
        ECore.NewTransform
    End If
    If GetTickCount - ReplaceAccept >= 5000 And ReplaceAccept > 0 Then ReplaceAccept = 0: ECore.NewTransform
End Sub
Public Sub DrawBackup()
    Page.Writes "备份清单", 381, 135, 18, argb(255, 35, 170, 242)
    
    For I = 1 To UBound(BackupList)
        If BackupList(I).UseTime = -1 Then
            Page.ShowLoading 600, 340, 128, 128, 3, argb(255, 35, 170, 242), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
            Exit Sub
        End If
    Next
    
    For I = BackupIndex * 6 + 1 To BackupIndex * 6 + 5 + 1
        If I > UBound(BackupList) Or I < 1 Then Exit For
        If I Mod 2 = 1 Then
            'Page.Paint 0, 380, 120 + 70 * (I - BackupIndex * 6) - 5 - 8, 560, 71, argb(255, 242, 242, 242)
        Else
            Page.Paint 0, 380, 120 + 70 * (I - BackupIndex * 6) - 5 - 8, 560, 71, argb(70, 242, 242, 255)
        End If
        If BackupList(I).Name = EmrPC.LastBranch Then
            Page.Paint 0, 380, 120 + 70 * (I - BackupIndex * 6) - 5 - 8, 7, 71, argb(255, 35, 170, 242)
        End If
        Page.Writes Left(BackupList(I).Name, 22) & IIf(Len(BackupList(I).Name) > 22, "...", "") & "  (" & BackupList(I).FileCount & " 个文件)", 392, 120 + 70 * (I - BackupIndex * 6) - 2, 18, argb(255, 255, 255, 255)
        Page.Writes BackupList(I).Date & " (" & BackupList(I).size & "字节)", 392, 120 + 70 * (I - BackupIndex * 6) - 2 + 25, 18, argb(255, 192, 192, 192)
        If Page.ShowColorButton(0, GW - 170, 117 + 70 * (I - BackupIndex * 6) + 8, 100, 30, IIf(BackupList(I).UseTime = 0, "还原", "确认 (" & Round(5 - (GetTickCount - BackupList(I).UseTime) / 1000) & "s)"), argb(255, 255, 255, 255), IIf(BackupList(I).UseTime = 0, argb(255, 35, 170, 242), argb(255, 254, 84, 57)), IIf(BackupList(I).UseTime = 0, argb(255, 75, 210, 255), argb(255, 255, 124, 97)), 40, size:=18) = 3 Then
            If BackupList(I).UseTime > 0 Then
                If BackupList(I).size = -1 Then
                    ShellExecuteA 0, "open", OPath & "\.emr\backup\", "", "", SW_SHOW
                    BackupList(I).UseTime = 0
                Else
                    BackupList(I).UseTime = -1
                    Dim Backup As EmrBackup, Cache As EmrBackup
                    Open OPath & "\.emr\" & IIf(BackupList(I).Name = "null", "", BackupList(I).Name) & ".emrbackup" For Binary As #1
                    Get #1, , Backup
                    Close #1
                    Cache.AFileHeader = Backup.AFileHeader
                    ReDim Cache.Files(UBound(Backup.Files))
                    For S = 1 To UBound(Backup.Files)
                        Open OPath & "\" & Backup.Files(S).path For Binary As #1
                        Put #1, , Backup.Files(S).Data
                        Close #1
                        Cache.Files(S).MD5Check = Backup.Files(S).MD5Check
                        Cache.Files(S).path = Backup.Files(S).path
                        ECore.Display: DoEvents
                    Next
                    EmrPC.FUpdate = True
                    EmrPC.LastBranch = BackupList(I).Name
                    Open OPath & "\.emerald" For Binary As #1
                    Put #1, , EmrPC
                    Close #1
                    Open OPath & "\.emr\.cache" For Binary As #1
                    Put #1, , Cache
                    Close #1
                    ReDim UFiles(0)
                    Call UpdateProject
                    BackupList(I).UseTime = 0
                    MsgBox "成功。", 64
                End If
                ECore.NewTransform
            Else
                BackupList(I).UseTime = GetTickCount
                ECore.NewTransform
            End If
        End If
        If GetTickCount - BackupList(I).UseTime >= 5000 And BackupList(I).UseTime > 0 Then BackupList(I).UseTime = 0: ECore.NewTransform
    Next
    
    Dim max As Long
    max = Int(UBound(BackupList) / 6) + 1
    Page.Writes BackupIndex + 1 & "/" & max, 600, 430 + 30 * 7 + 5, 18, argb(255, 242, 242, 242), 120, 20, align:=StringAlignmentCenter
    If BackupIndex + 1 > 1 Then
        If Page.ShowColorButton(1, 600, 430 + 30 * 7 + 5, 20, 20, "<", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=16) = 3 Then
            BackupIndex = BackupIndex - 1
            ECore.NewTransform
        End If
    End If
    If BackupIndex + 1 < max Then
        If Page.ShowColorButton(1, 700, 430 + 30 * 7 + 5, 20, 20, ">", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=16) = 3 Then
            BackupIndex = BackupIndex + 1
            ECore.NewTransform
        End If
    End If
    If BackupIndex + 1 > max Then BackupIndex = max - 1
End Sub
Public Sub DrawPackage()
    Page.Writes "打包", 381, 135, 18, argb(255, 35, 170, 242)
    Page.Writes "将你的游戏打包为安装包。", 381, 175, 18, argb(255, 242, 242, 242)
    
    If PackState = 1 Then
        Page.Writes "正在制作安装包...", 381, 205, 18, argb(255, 242, 242, 242)
        Page.Writes PackText, 381, 235, 18, argb(255, 242, 242, 242)
        Page.ShowLoading 600, 410, 128, 128, 3, argb(255, 35, 170, 242), argb(255, 0, 176, 240), argb(255, 64, 218, 208)
        Exit Sub
    End If
        
    Dim CanDo As Boolean
    CanDo = True
    If Dir(OPath & "\app.exe") = "" Then
        CanDo = False
        Page.Writes "[待处理]     目录下需要一个名为app.exe的游戏主程序。", 381, 255, 18, argb(255, 254, 84, 57)
    Else
        Page.Writes "[已处理]     目录下需要一个名为app.exe的游戏主程序。", 381, 255, 18, argb(255, 35, 170, 242)
    End If
    If Dir(OPath & "\app.png") = "" Then
        CanDo = False
        Page.Writes "[待处理]     目录下需要一张名为app.png的游戏图标。", 381, 285, 18, argb(255, 254, 84, 57)
    Else
        Page.Writes "[已处理]     目录下需要一张名为app.png的游戏图标。", 381, 285, 18, argb(255, 35, 170, 242)
    End If
    If CanDo = False Then
        Page.Writes "我们无法继续制作安装包。", 381, 205, 18, argb(255, 242, 242, 242)
    Else
        Page.Writes "一切准备就绪！", 381, 205, 18, argb(255, 242, 242, 242)
        If Page.ShowColorButton(0, GW - 220, GH - 112, 150, 40, "制作", argb(255, 255, 255, 255), argb(255, 35, 170, 242), argb(255, 75, 210, 255), 40, size:=18) = 3 Then
            PackState = 1: PackText = "正在准备..."
            Call BuildInstaller
            PackState = 0
            MsgBox "制作成功！", 64
        End If
    End If
End Sub
Public Sub DrawSettings()
    Dim Data As String
    Page.Writes "调试工具", 381, 135, 18, argb(255, 35, 170, 242)
    Data = ESave.GetData("DebugMode")
    Page.ShowColorCheckBox Data, 384, 175, 200, 20, "在屏幕顶部显示调试栏", argb(255, 242, 242, 242), argb(255, 35, 170, 242), argb(255, 242, 242, 242), 18
    If Data <> ESave.GetData("DebugMode") Then ESave.PutData "DebugMode", IIf(Data = True, 1, 0)
    
    Data = ESave.GetData("DisableLOGO")
    Page.ShowColorCheckBox Data, 384, 205, 200, 20, "在所有资源加载完成后跳过LOGO动画", argb(255, 242, 242, 242), argb(255, 35, 170, 242), argb(255, 242, 242, 242), 18
    If Data <> ESave.GetData("DisableLOGO") Then ESave.PutData "DisableLOGO", IIf(Data = True, 1, 0)
    
    Data = ESave.GetData("HideLOGO")
    Page.ShowColorCheckBox Data, 384, 235, 200, 20, "不要显示LOGO动画", argb(255, 242, 242, 242), argb(255, 35, 170, 242), argb(255, 242, 242, 242), 18
    If Data <> ESave.GetData("HideLOGO") Then ESave.PutData "HideLOGO", IIf(Data = True, 1, 0)
    
     Data = ESave.GetData("HideSuggest")
    Page.ShowColorCheckBox Data, 384, 265, 200, 20, "关闭建议中心", argb(255, 242, 242, 242), argb(255, 35, 170, 242), argb(255, 242, 242, 242), 18
    If Data <> ESave.GetData("HideSuggest") Then ESave.PutData "HideSuggest", IIf(Data = True, 1, 0)
    
    Data = (Dir(VBIDEPath & "Template\Forms\Emerald Window.frm") <> "")
    Page.ShowColorCheckBox Data, 384, 295, 200, 20, "VB工程示例文件", argb(255, 242, 242, 242), argb(255, 35, 170, 242), argb(255, 242, 242, 242), 18
    If Data <> (Dir(VBIDEPath & "Template\Forms\Emerald Window.frm") <> "") Then
        On Error GoTo sth
        If Dir(VBIDEPath & "Template", vbDirectory) = "" Then MkDir VBIDEPath & "Template"
        If Data = False Then
            Kill VBIDEPath & "Template\Forms\Emerald Window.frm"
            Kill VBIDEPath & "Template\Classes\Emerald Page.cls"
        Else
            FileCopy App.path & "\Example\Emerald Window.frm", VBIDEPath & "Template\Forms\Emerald Window.frm"
            FileCopy App.path & "\Example\Emerald Page.cls", VBIDEPath & "Template\Classes\Emerald Page.cls"
        End If
sth:
        If Err.Number <> 0 Then
            MsgBox "啊，失败了。" & vbCrLf & Err.Description, 16
            Err.Clear
        End If
    End If
    
    Page.Writes "更新计划", 381, 345, 18, argb(255, 35, 170, 242)
    Page.Writes "检查更新间隔 (天) : ", 381, 380, 18, argb(255, 242, 242, 242)
    Data = ESave.GetData("UpdateCheckInterval")
    Page.ShowEdit Data, 0, 581, 380 - 5 + 1, 280, 30, argb(255, 35, 170, 242), argb(0, 255, 255, 255), argb(50, 242, 242, 242), argb(255, 35, 170, 242), 50, size:=18
    If Data <> ESave.GetData("UpdateCheckInterval") Then ESave.PutData "UpdateCheckInterval", Val(Data)
    Page.Writes "连接超时时间 (毫秒) : ", 381, 410, 18, argb(255, 242, 242, 242)
    Dim data2 As String
    data2 = ESave.GetData("UpdateTimeOut")
    Page.ShowEdit data2, 0, 581, 410 - 5 + 1, 280, 30, argb(255, 35, 170, 242), argb(0, 255, 255, 255), argb(50, 242, 242, 242), argb(255, 35, 170, 242), 50, size:=18
    If data2 <> ESave.GetData("UpdateTimeOut") Then ESave.PutData "UpdateTimeOut", Val(data2)

    Dim unAble As Boolean
    If ESave.GetData("UpdateTime") <> "" Then
        On Error Resume Next
        If Not (Now - CDate(ESave.GetData("UpdateTime")) >= UpdateCheckInterval) Then
            m = CheckMouse(381, 445, 170, 20)
            Page.Writes "重置更新 ...", 381, 445, 18, IIf(m = 0, argb(255, 242, 242, 242), argb(255, 35, 170, 242)), 200, 20, style:=FontStyleUnderline
            If m = 3 Then
                ESave.PutData "UpdateAble", 0
                ESave.PutData "UpdateTime", ""
                MsgBox "重置更新成功。", 64
                ECore.NewTransform transFadeIn, 1000
            End If
        Else
            unAble = True
        End If
        Err.Clear
    Else
        unAble = True
    End If
    
    If unAble Then
        m = CheckMouse(381, 445, 250, 20)
        Page.Writes "现在无法重置更新。", 381, 445, 18, argb(255, 148, 148, 148), 250, 20
        m = IIf(m = 0, MousePointerConstants.vbDefault, MousePointerConstants.vbNoDrop)
        If Screen.MousePointer <> m Then Screen.MousePointer = m
    End If
    
    Page.Writes "资源管理器菜单", 381, 505, 18, argb(255, 35, 170, 242)
    Page.Writes "将Emerald固定在资源管理器菜单上使创建工程更加便捷。", 381, 545, 18, argb(255, 242, 242, 242)
    Page.Writes "你将在你的资源管理器中找到它 (右键背景) .", 381, 575, 18, argb(255, 242, 242, 242)
    Data = IIf(IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\") = True, 1, 0)
    Page.ShowColorCheckBox Data, 384, 615, 200, 20, "Emerald菜单项", argb(255, 242, 242, 242), IIf(CMError = "", argb(255, 35, 170, 242), argb(255, 242, 242, 242)), IIf(CMError = "", argb(255, 242, 242, 242), argb(255, 242, 242, 242)), 18
    If Data <> IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\") And CMError = "" Then
        CMError = OperContentMenu(IIf(Not CBool(Data), True, False))
        If CMError <> "" Then ECore.NewTransform transFadeIn, 300
    End If
    
    If CMError <> "" Then
        Page.Writes " !    " & CMError, 381, 655, 18, argb(255, 254, 84, 57), 600, 20
        If CheckMouse2 = mMouseUp Then MsgBox "详细信息:" & vbCrLf & "Builder在尝试" & IIf(CBool(Data), "移除", "添加") & "菜单项时出错了。" & vbCrLf & vbCrLf & "详细信息:" & vbCrLf & CMError, 64
    End If
End Sub
Public Sub Update()
    '绘制事件
    
    'MsgBox 1 / 0
    
    Page.Clear 0
    
    If Cliped = 1 Then
        GdipSetClipRect Page.GG, 38 * Scales, 38 * Scales, RGW - 38 * 2 * Scales, RGH - 38 * 2 * Scales, CombineModeReplace
        Dim Rgn As Long
        Rgn = CreateRectRgn(38 * Scales, 80 * Scales, RGW - 38 * Scales, RGH - 38 * Scales)
        SelectClipRgn GCore.GDC, Rgn
        SelectClipRgn ECore.CDC, Rgn
        Cliped = 2
    End If
    
    Page.Clear argb(0, 0, 0, 0)
    Page.DrawImage "background" & IIf((ESave.GetData("Emerald 3.1") = "") Or (ReadProject = False And OPath <> "" And Dir(OPath & "\.emerald") = ""), "2", "") & ".png", 0, 0, Alpha:=1
    If Cliped = 0 Then Cliped = 1
    
    'Page.PaintLine 342, 32 + 24, 342, GH - 40 - 24 + 1, argb(255, 235, 235, 235), 1
    
    Dim m As Long, HY As Long
    
    m = CheckMouse(48, 52, 48, 48)
    
    HY = 128
    For I = 0 To 5
        If I = 2 Then GoTo SkipThis
        m = CheckMouse(0, HY, 313, 41)
        If I = PageIndex Then
            Page.Paint 0, 38 * Scales, HY, (313 - 38) * Scales, 41, argb(30, 255, 255, 255), 100
            Page.Paint 0, 38 * Scales, HY, (45 - 38) * Scales, 41, argb(255, 35, 170, 242), 100
        Else
            If m Then
                Page.Paint 0, 38 * Scales, HY, (313 - 38) * Scales, 41, argb(60, 255, 255, 255), 100
            Else
                Page.Paint 0, 38 * Scales, HY, (313 - 38) * Scales, 41, argb(0, 255, 255, 255), 100
            End If
        End If
        If m = mMouseUp Then
            If OPath = "" And I <= 3 Then
                MsgBox "请在你的Emerald工程文件夹运行Builder。", 64
            Else
                PageIndex = I
                'ECore.NewTransform transFadeIn, 1000
            End If
        End If
        HY = HY + 50
SkipThis:
    Next
    
    Page.DrawImage "project.png", 60, 135, Alpha:=1
    Page.Writes "工程", 110, 135, 18, argb(255, 255, 255, 255)
    
    Page.DrawImage "backup.png", 60, 135 + 50, Alpha:=1
    Page.Writes "备份", 110, 135 + 50, 18, argb(255, 255, 255, 255)
    
    Page.DrawImage "package.png", 60, 135 + 50 * 2, Alpha:=1
    Page.Writes "打包", 110, 135 + 50 * 2, 18, argb(255, 255, 255, 255)
    
    Page.DrawImage "settings.png", 60, 135 + 50 * 3, Alpha:=1
    Page.Writes "设置", 110, 135 + 50 * 3, 18, argb(255, 255, 255, 255)
    
    Page.DrawImage "information.png", 60, 135 + 50 * 4, Alpha:=1
    Page.Writes "关于", 110, 135 + 50 * 4, 18, argb(255, 255, 255, 255)
    
    Select Case PageIndex
        Case 0: DrawProject
        Case 1: DrawBackup
        Case 3: DrawPackage
        Case 4: DrawSettings
        Case 5: DrawAbout
    End Select
     
    If ESave.GetData("Emerald 3.1") = "" And Cliped = 2 Then
        ECore.ActivePage = "ToNewPage"
        Exit Sub
    End If
    If ReadProject = False And OPath <> "" And Cliped = 2 Then
        ReadProject = True
        If Dir(OPath & "\.emerald") = "" Then
            ECore.ActivePage = "ToNewPage"
            ToNewPage.PageIndex = 5
            Exit Sub
        End If
        Open OPath & "\.emerald" For Binary As #1
        Get #1, , EmrPC
        Close #1
        If EmrPC.AFileHeader <> "Emerald Project Config File" Then
            ECore.ActivePage = "ToNewPage"
            ToNewPage.PageIndex = 4
        End If
    End If
    
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub InitProject()
    ReDim UFiles(0)
    ReDim BackupList(0)
    BackupSwitch = True
    
    If OPath = App.path Then
        OPath = "": PageIndex = 4
        MsgBox "请在其他地方运行Builder。", 48
        Exit Sub
    End If
    
    If Dir(OPath & "\.emerald") = "" Then Exit Sub
    Call UpdateProject
    Call ScanBackup
End Sub
Public Sub Enter()
    '页面进入事件
    If OPath = "" Then PageIndex = 4
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
    If PackPos = -1 Then
        Page.Res.NewImages App.path & "\assets", 0.75
        Page.Res.NewImages App.path & "\assets\icons", 24, 24
    End If
    StartTime = GetTickCount
    
    '创建页面
    ECore.Add Page, "WelcomePage"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
        Page.NewAnimation False, "FadeIn", 2, 500
        Page.NewAnimation False, "FadeIn", 2, 500
    '===============================================
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub

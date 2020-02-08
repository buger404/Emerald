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
    Public BackupC As Integer
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
    BackupC = 1
    
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
                If .Name = EmrPC.LastBranch Then BackupC = UBound(BackupList)
                .size = FileLen(OPath & "\.emr\" & File)
            End With
        End If
        File = Dir()
    Loop
    
    If Dir(OPath & "\.emr\backup", vbDirectory) <> "" Then
        ReDim Preserve BackupList(UBound(BackupList) + 1)
        With BackupList(UBound(BackupList))
            .Date = "使用远古版本Builder建立的备份"
            .FileCount = -1
            .Name = "远古"
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

    Page.Writes "框架版本", 85, 190, 18, argb(255, 0, 0, 0)
    Page.Writes "20" & Mid(Version, 1, 2) & "." & Val(Mid(Version, 3, 2)) & "." & Val(Mid(Version, 5, 2)) & " 第" & Val(Mid(Version, 7, 2)) & "次版本", 200, 190, 18, argb(200, 32, 32, 36)
    Page.Writes "制作", 85, 230, 18, argb(255, 0, 0, 0)
    Page.Writes "Error404 (QQ 1361778219)", 200, 230, 18, argb(200, 32, 32, 36)
    Page.Writes "支持", 85, 270, 18, argb(255, 0, 0, 0)
    Page.Writes "Inter.Net，南笙", 200, 270, 18, argb(200, 32, 32, 36)
    Page.Writes "更新", 85, 310, 18, argb(255, 0, 0, 0)
    If NewVersion = 0 Then
        Page.ShowLoading 200, 310 + 2, 20, 20, 3, argb(255, 0, 120, 215), argb(255, 0, 176, 240), argb(160, 64, 218, 208)
        Page.Writes "正在检查更新...", 230, 310, 18, argb(200, 32, 32, 36)
    ElseIf NewVersion = 3 Then
        Page.Writes "连接超时", 200, 310, 18, argb(200, 254, 84, 57)
    ElseIf NewVersion > Version Then
        Page.Writes "更新可用", 200, 310, 18, argb(255, 35, 170, 242), 200, 22, style:=FontStyleUnderline
        If CheckMouse2 = mMouseUp Then ShellExecuteA 0, "open", "https://github.com/buger404/Emerald/releases", "", "", SW_SHOW
    Else
        Page.Writes "官方最新版", 200, 310, 18, argb(200, 32, 32, 36)
    End If
    
    Page.Writes "使用帮助", 85, 430, 18, argb(255, 0, 0, 0)
    m = CheckMouse(200, 430, 195, 20)
    Page.Writes "Github : Emerald Wiki", 200, 430, 18, IIf(m = 0, argb(200, 32, 32, 36), argb(255, 60, 205, 195)), 200, 25, style:=FontStyleUnderline
    If m = 3 Then ShellExecuteA 0, "open", "https://github.com/buger404/Emerald/wiki", "", "", SW_SHOW
    m = IIf(m = 0, MousePointerConstants.vbDefault, MousePointerConstants.vbArrowQuestion)
    If Screen.MousePointer <> m Then Screen.MousePointer = m
    
    Dim Data As String
    Page.Writes "检查间隔", 85, 350, 18, argb(255, 0, 0, 0)
    Data = ESave.GetData("UpdateCheckInterval") & "天"
    Page.ShowEdit Data, 0, 200, 350 - 1, 280, 30, argb(200, 32, 32, 36), argb(0, 16, 16, 24), argb(255, 242, 242, 242), argb(255, 62, 206, 196), size:=18
    Data = Replace(Data, "天", "")
    If Data <> ESave.GetData("UpdateCheckInterval") Then ESave.PutData "UpdateCheckInterval", Val(Data)
    Page.Writes "超时时间", 85, 390, 18, argb(255, 0, 0, 0)
    Dim data2 As String
    data2 = ESave.GetData("UpdateTimeOut") & "ms"
    Page.ShowEdit data2, 0, 200, 390 - 1, 280, 30, argb(200, 32, 32, 36), argb(0, 16, 16, 24), argb(255, 242, 242, 242), argb(255, 62, 206, 196), size:=18
    data2 = Replace(data2, "ms", "")
    If data2 <> ESave.GetData("UpdateTimeOut") Then ESave.PutData "UpdateTimeOut", Val(data2)

End Sub
Public Sub DrawProject()
    Page.Writes "项目名称", 85, 190, 18, argb(255, 0, 0, 0)
    Page.Writes EmrPC.Name, 200, 190, 18, argb(200, 32, 32, 36)
    Page.Writes "开发者", 85, 230, 18, argb(255, 0, 0, 0)
    If Page.ShowEdit(EmrPC.Maker, 0, 200, 230 - 1, 280, 30, argb(200, 32, 32, 36), argb(0, 16, 16, 24), argb(255, 242, 242, 242), argb(255, 62, 206, 196), size:=18) = 4 Then
        Open OPath & "\.emerald" For Binary As #1
        Put #1, , EmrPC
        Close #1
    End If
    
    Dim m As Long
    
    Page.Writes "框架版本", 85, 270, 18, IIf(UBound(UFiles) > 0, argb(255, 254, 84, 57), argb(255, 0, 0, 0))
    Page.Writes "V" & EmrPC.Version & IIf(UBound(UFiles) > 0, " -> V" & Version, ""), 200, 270, 18, argb(200, 32, 32, 36)
    
    If UBound(UFiles) = 0 Then Exit Sub
    If ReplaceAccept = -1 Then
        Page.ShowLoading GW / 2 - 24, 355, 48, 48, 3, argb(255, 0, 120, 215), argb(255, 0, 176, 240), argb(160, 64, 218, 208)
        Exit Sub
    End If
    
    Dim UColor(3) As Long, UText(3) As String
    UColor(1) = argb(255, 67, 219, 210): UText(1) = "新增"
    UColor(2) = argb(255, 35, 170, 242): UText(2) = "更新"
    UColor(3) = argb(255, 254, 84, 57): UText(3) = "替换"
    
    For I = FileIndex * 3 + 1 To FileIndex * 3 + 2 + 1
        If I > UBound(UFiles) Or I < 1 Then Exit For
        Page.Paint 1, 85 + 3, 310 + 30 * (I - FileIndex * 3) - 30 + 3, 15, 15, IIf(UFiles(I).ReplaceMark, UColor(UFiles(I).Kind), argb(255, 200, 200, 200))
        If CheckMouse2 = mMouseUp Then UFiles(I).ReplaceMark = Not UFiles(I).ReplaceMark
        Page.Writes UFiles(I).Name, 200, 310 + 30 * (I - FileIndex * 3) - 30, 18, IIf(UFiles(I).ReplaceMark, argb(200, 16, 16, 24), argb(255, 160, 160, 160))
        Page.Writes UText(UFiles(I).Kind), 110, 310 + 30 * (I - FileIndex * 3) - 30, 18, IIf(UFiles(I).ReplaceMark, UColor(UFiles(I).Kind), argb(255, 160, 160, 160))
        Page.Writes IIf(UFiles(I).IncraseByte > 0, "+", "") & UFiles(I).IncraseByte & " 字节", GW - 80, 310 + 30 * (I - FileIndex * 3) - 30, 18, argb(255, 96, 96, 100), align:=StringAlignmentFar
    Next
    
    Dim max As Long
    max = Int(UBound(UFiles) / 3) + 1
    
    If FileIndex + 1 > 1 Then
        If Page.ShowColorButton(1, 40, 338, 40, 40, "<", argb(255, 16, 16, 24), argb(0, 35, 170, 242), argb(128, 65, 207, 197), 40, size:=16) = 3 Then
            FileIndex = FileIndex - 1
        End If
    End If
    If FileIndex + 1 < max Then
        If Page.ShowColorButton(1, GW - 80, 338, 40, 40, ">", argb(255, 16, 16, 24), argb(0, 35, 170, 242), argb(128, 65, 207, 197), 40, size:=16) = 3 Then
            FileIndex = FileIndex + 1
        End If
    End If
    If FileIndex + 1 > max Then FileIndex = max - 1
    
    BackupSwitch = True
    
    If Page.ShowColorButton(2, GW / 2 - 200, GH - 110, 400, 40, IIf(ReplaceAccept = 0, "替换代码", "确认替换 (" & Round(5 - (GetTickCount - ReplaceAccept) / 1000) & "s)"), argb(255, 255, 255, 255), IIf(ReplaceAccept = 0, argb(255, 103, 226, 217), argb(255, 254, 84, 57)), IIf(ReplaceAccept = 0, argb(255, 60, 205, 195), argb(255, 255, 124, 97)), 40, size:=18, style:=FontStyleBold) = 3 Then
        If ReplaceAccept > 0 Then
            ReplaceAccept = -1
            Call ReplaceProject
            Call ScanBackup
            ReplaceAccept = 0
            MsgBox "Emerald成功替换了旧版本代码。" & IIf(BackupSwitch, vbCrLf & "并成功创建了旧版本的备份 : " & OPath & "\.emr\" & EmrPC.LastBranch & ".emrbackup", ""), 64
        Else
            ReplaceAccept = GetTickCount
        End If
    End If
    If GetTickCount - ReplaceAccept >= 5000 And ReplaceAccept > 0 Then ReplaceAccept = 0
End Sub
Public Sub DrawBackup()

    For I = 1 To UBound(BackupList)
        If BackupList(I).UseTime = -1 Then
            Page.ShowLoading GW / 2 - 24, 355, 48, 48, 3, argb(255, 0, 120, 215), argb(255, 0, 176, 240), argb(160, 64, 218, 208)
            Exit Sub
        End If
    Next
    
    For I = BackupIndex * 3 + 1 To BackupIndex * 3 + 2 + 1
        If I > UBound(BackupList) Or I < 1 Then Exit For

        Page.Writes "备份" & I, 85, 190 + 70 * (I - BackupIndex * 3) - 70, 18, argb(IIf(BackupC = I, 255, 120), 0, 0, 0), style:=IIf(BackupC = I, FontStyleBold, FontStyleRegular)
        If BackupList(I).FileCount = -1 Then
            Page.Writes BackupList(I).Date, 200, 190 + 70 * (I - BackupIndex * 3) - 70, 18, argb(IIf(BackupC = I, 200, 120), 16, 16, 29)
            Page.Writes "过于老旧的备份，无法取得信息", 200, 190 + 70 * (I - BackupIndex * 3) - 70 + 25, 18, argb(IIf(BackupC = I, 200, 120), 16, 16, 29)
        Else
            Page.Writes "日期 " & BackupList(I).Date, 200, 190 + 70 * (I - BackupIndex * 3) - 70, 18, argb(IIf(BackupC = I, 200, 120), 16, 16, 29)
            Page.Writes "共" & BackupList(I).FileCount & "个文件，总计" & Int(BackupList(I).size / 1024) & "kb", 200, 190 + 70 * (I - BackupIndex * 3) - 70 + 25, 18, argb(IIf(BackupC = I, 200, 120), 16, 16, 29)
        End If
        
        If CheckMouse(85, 190 + 70 * (I - BackupIndex * 3) - 70, GW - 85 * 2, 70) = mMouseUp Then BackupC = I
    Next
    
    I = BackupC
   If Page.ShowColorButton(2, GW / 2 - 200, GH - 110, 400, 40, IIf(BackupList(I).UseTime = 0, "恢复备份" & I, "确认恢复备份" & I & " (" & Round(5 - (GetTickCount - BackupList(I).UseTime) / 1000) & "s)"), argb(255, 255, 255, 255), IIf(BackupList(I).UseTime = 0, argb(255, 103, 226, 217), argb(255, 254, 84, 57)), IIf(BackupList(I).UseTime = 0, argb(255, 60, 205, 195), argb(255, 255, 124, 97)), 40, size:=18, style:=FontStyleBold) = 3 Then
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
        End If
    End If
    If GetTickCount - BackupList(I).UseTime >= 5000 And BackupList(I).UseTime > 0 Then BackupList(I).UseTime = 0
    
    Dim max As Long
    max = Int(UBound(BackupList) / 3) + 1
    If BackupIndex + 1 > 1 Then
        If Page.ShowColorButton(1, 40, 268, 40, 40, "<", argb(255, 16, 16, 24), argb(0, 35, 170, 242), argb(128, 65, 207, 197), 40, size:=16) = 3 Then
            BackupIndex = BackupIndex - 1
        End If
    End If
    If BackupIndex + 1 < max Then
        If Page.ShowColorButton(1, GW - 80, 268, 40, 40, ">", argb(255, 16, 16, 24), argb(0, 35, 170, 242), argb(128, 65, 207, 197), 40, size:=16) = 3 Then
            BackupIndex = BackupIndex + 1
        End If
    End If
    If BackupIndex + 1 > max Then BackupIndex = max - 1
End Sub
Public Sub DrawPackage()

    If PackState = 1 Then
        Page.Writes "进度", 85, 190, 18, argb(255, 0, 0, 0)
        Page.Writes PackText, 200, 190, 18, argb(200, 32, 32, 36)
        Page.ShowLoading GW / 2 - 24, 355, 48, 48, 3, argb(255, 0, 120, 215), argb(255, 0, 176, 240), argb(160, 64, 218, 208)
        Exit Sub
    End If
        
    Dim CanDo As Boolean
    CanDo = True
    If Dir(OPath & "\app.exe") = "" Then
        CanDo = False
        Page.Writes "需要", 85, 190, 18, argb(255, 160, 160, 160)
        Page.Writes "主程序 'app.exe'", 200, 190, 18, argb(200, 160, 160, 160)
    Else
        Page.Writes "需要", 85, 190, 18, argb(255, 0, 0, 0)
        Page.Writes "主程序 'app.exe'", 200, 190, 18, argb(200, 16, 16, 32)
    End If
    
    If Dir(OPath & "\app.png") = "" Then
        CanDo = False
        Page.Writes "需要", 85, 230, 18, argb(255, 160, 160, 160)
        Page.Writes "图标 'app.png'", 200, 230, 18, argb(200, 160, 160, 160)
    Else
        Page.Writes "需要", 85, 230, 18, argb(255, 0, 0, 0)
        Page.Writes "图标 'app.png'", 200, 230, 18, argb(200, 16, 16, 32)
    End If
    If CanDo Then
        If Page.ShowColorButton(2, GW / 2 - 200, GH - 110, 400, 40, "制作安装程序", argb(255, 255, 255, 255), argb(255, 105, 227, 218), argb(255, 59, 205, 195), 40, size:=18, style:=FontStyleBold) = 3 Then
            PackState = 1: PackText = "正在准备..."
            Call BuildInstaller
            PackState = 0
            MsgBox "制作成功！", 64
        End If
    End If
End Sub
Public Sub DrawSettings()
    Dim Data As String

    Data = ESave.GetData("DebugMode")
    Page.ShowColorCheckBox Data, 85, 190, 200, 20, "调试时在屏幕顶部显示调试栏", argb(255, 200, 200, 200), argb(255, 60, 205, 195), argb(255, 32, 32, 36), 18
    If Data <> ESave.GetData("DebugMode") Then ESave.PutData "DebugMode", IIf(Data = True, 1, 0)
    
    Data = ESave.GetData("DisableLOGO")
    Page.ShowColorCheckBox Data, 85, 230, 200, 20, "在所有资源加载完成后跳过加载动画", argb(255, 200, 200, 200), argb(255, 60, 205, 195), argb(255, 32, 32, 36), 18
    If Data <> ESave.GetData("DisableLOGO") Then ESave.PutData "DisableLOGO", IIf(Data = True, 1, 0)
    
    Data = ESave.GetData("HideLOGO")
    Page.ShowColorCheckBox Data, 85, 270, 200, 20, "禁用加载动画", argb(255, 200, 200, 200), argb(255, 60, 205, 195), argb(255, 32, 32, 36), 18
    If Data <> ESave.GetData("HideLOGO") Then ESave.PutData "HideLOGO", IIf(Data = True, 1, 0)
    
     Data = ESave.GetData("HideSuggest")
    Page.ShowColorCheckBox Data, 85, 310, 200, 20, "关闭建议中心", argb(255, 200, 200, 200), argb(255, 60, 205, 195), argb(255, 32, 32, 36), 18
    If Data <> ESave.GetData("HideSuggest") Then ESave.PutData "HideSuggest", IIf(Data = True, 1, 0)
    
    Data = (Dir(VBIDEPath & "Template\Forms\Emerald Window.frm") <> "")
    Page.ShowColorCheckBox Data, 85, 350, 200, 20, "使用Builder提供的VB工程示例文件", argb(255, 200, 200, 200), argb(255, 60, 205, 195), argb(255, 32, 32, 36), 18
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

    Data = IIf(IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\") = True, 1, 0)
    Page.ShowColorCheckBox Data, 85, 390, 200, 20, "资源管理器菜单", argb(255, 32, 32, 36), IIf(CMError = "", argb(255, 60, 205, 195), argb(255, 32, 32, 36)), IIf(CMError = "", argb(255, 32, 32, 36), argb(255, 32, 32, 36)), 18
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
        'GdipSetClipRect Page.GG, 20 * Scales, 20 * Scales, RGW - 20 * 2 * Scales, RGH - 20 * 2 * Scales, CombineModeReplace
        'Dim Rgn As Long
        'Rgn = CreateRectRgn(20 * Scales, 80 * Scales, RGW - 20 * Scales, RGH - 20 * Scales)
        'SelectClipRgn GCore.GDC, Rgn
        'SelectClipRgn ECore.CDC, Rgn
        Cliped = 2
    End If
    
    Page.Clear argb(0, 0, 0, 0)
    Page.DrawImage "background.png", 0, 0, alpha:=1
    If Cliped = 0 Then Cliped = 1
    
    'Page.PaintLine 342, 32 + 24, 342, GH - 40 - 24 + 1, argb(255, 235, 235, 235), 1
    
    Dim m As Long, HY As Long, Menus(5) As String
    Menus(0) = "属性"
    Menus(1) = "备份"
    Menus(3) = "发布"
    Menus(4) = "系统"
    Menus(5) = "关于"
    
    m = CheckMouse(48, 52, 48, 48)
    
    HY = GW / 2 - 5 * 80 / 2
    
    For I = 0 To 5
        If I = 2 Then GoTo SkipThis
        m = CheckMouse(HY, 135, 80, 20)
        Page.Writes Menus(I), HY + 40, 135, 18, argb(IIf(m > 0 Or I = PageIndex, 255, 160), 16, 16, 24), align:=StringAlignmentCenter, style:=IIf(I = PageIndex, FontStyleBold, FontStyleRegular)
        If I = PageIndex Then
            Page.Paint 0, HY + 10, 165, 60, 6, argb(255, 60, 205, 195), 100
        End If
        If m = 3 Then
            If OPath = "" And I <= 3 Then
                DrawF.X = HY: DrawF.y = 135: DrawF.Width = 80: DrawF.Height = 20
                MsgBox "请在你的Emerald工程文件夹运行Builder。", 48
            Else
                PageIndex = I
            End If
        End If
        HY = HY + 80
SkipThis:
    Next
    
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
        Page.Res.NewImages App.path & "\assets", 0.8
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

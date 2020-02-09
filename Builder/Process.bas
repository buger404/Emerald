Attribute VB_Name = "Process"
'Emerald 相关代码

Public InstalledPath As String, IsUpdate As Boolean

Public Tasks() As String
Public CmdMark As String, SetupErr As Long, Repaired As Boolean
Public AppInfo() As String
Public Cmd As String
Public Abouting As Boolean
Public Function TestFile(path As String, IncludeText As String) As Boolean
    Dim temp As String
    If Dir(path) = "" Then Exit Function
    Open path For Input As #1
    Do While Not EOF(1)
        Line Input #1, temp
        If InStr(temp, IncludeText) > 0 Then TestFile = True: Exit Do
    Loop
    Close #1
End Function
Public Sub CheckUpdate2()
    On Error GoTo ErrHandle
    
    Dim WSHShell As Object, temp As String
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\version")
    IsUpdate = (Val(temp) <> Version)
    
ErrHandle:
    
End Sub
Public Sub GetInstalledPath()
    On Error GoTo ErrHandle
    
    Dim WSHShell As Object, temp As String
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\icon")
    InstalledPath = Replace(temp, """", "")
    
ErrHandle:
    
End Sub
Public Function CheckFileName(Name As String) As Boolean
    CheckFileName = ((InStr(Name, "*") Or InStr(Name, "\") Or InStr(Name, "/") Or InStr(Name, ":") Or InStr(Name, "?") Or InStr(Name, """") Or InStr(Name, "<") Or InStr(Name, ">") Or InStr(Name, "|") Or InStr(Name, " ") Or InStr(Name, "!") Or InStr(Name, "-") Or InStr(Name, "+") Or InStr(Name, "#") Or InStr(Name, "@") Or InStr(Name, "$") Or InStr(Name, "^") Or InStr(Name, "&") Or InStr(Name, "(") Or InStr(Name, ")")) = 0)
    Dim t As String
    If Name <> "" Then t = Left(Name, 1)
    CheckFileName = CheckFileName And (Trim(Str(Val(t))) <> t)
End Function
Sub Uninstall()
    'If Dialog("卸载", "Emerald Builder 已经安装，你希望删除它吗？", "卸载", "手滑") <> 1 Then End
    On Error Resume Next
    
    SetupPage.SetupInfo = "正在创建：WScript.Shell对象"
    SetupPage.Progress = 0.1
    Call FakeSleep
    
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    SetupPage.SetupInfo = "正在删除：资源管理器背景菜单项"
    SetupPage.Progress = 0.4
    Call FakeSleep
    
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\version"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\"
    
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\"
    
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\icon"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\command\"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\"
    
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\icon"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\command\"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\"
    
    SetupPage.SetupInfo = "正在删除：软件信息"
    SetupPage.Progress = 0.7
    Call FakeSleep
    
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayIcon"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayName"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayVersion"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\Publisher"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\URLInfoAbout"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\UninstallString"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\InstallLocation"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\"
    
    SetupPage.SetupInfo = "正在删除：Visual Basic 6 模板文件 (1/2)"
    SetupPage.Progress = 0.8
    Call FakeSleep
    
    Kill VBIDEPath & "Template\Forms\Emerald 游戏窗口.frm"
    
    SetupPage.SetupInfo = "正在删除：Visual Basic 6 模板文件 (2/2)"
    SetupPage.Progress = 0.9
    Call FakeSleep
    
    Kill VBIDEPath & "Template\Classes\Emerald 页面.cls"
    
    SetupPage.SetupInfo = "收尾"
    SetupPage.Progress = 1
    
    SetupErr = Err.Number
End Sub
Sub FakeSleep(Optional Counts As Long = 10)
    For I = 1 To Counts
        Sleep 10: DoEvents
        ECore.Display
    Next
End Sub
Sub Setup()
    On Error Resume Next
    
    Dim exeP As String
    exeP = """" & App.path & "\Builder.exe" & """"
    
    SetupPage.SetupInfo = "正在创建：WScript.Shell对象"
    SetupPage.Progress = 0.1
    Set WSHShell = PoolCreateObject("WScript.Shell")

    Call FakeSleep

    SetupPage.SetupInfo = "正在注册：资源管理器背景菜单项"
    SetupPage.Progress = 0.3
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\", "在此处创建/更新Emerald工程"
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon", exeP
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\version", Version
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\", exeP & " ""%v"""
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\", "在此处创建/更新Emerald工程"
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon", exeP
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\", exeP & " ""%v"""
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\", "制作该Emerald工程的安装包"
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\icon", exeP
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\command\", exeP & " p""%v"""
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\", "制作该Emerald工程的安装包"
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\icon", exeP
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\command\", exeP & " p""%v"""
    
    
    Call FakeSleep
    
    SetupPage.SetupInfo = "正在注册：软件信息"
    SetupPage.Progress = 0.6
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayIcon", exeP
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayName", "Emerald"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayVersion", "Indev " & Version
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\Publisher", "Error 404"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\InstallLocation", App.path
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\URLInfoAbout", "http://red-error404.github.io/233"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\UninstallString", exeP & " ""-uninstall"""
    
    Call FakeSleep
    
    SetupPage.SetupInfo = "正在复制：Visual Basic 6 模板文件（1/2）"
    SetupPage.Progress = 0.8
    FileCopy App.path & "\Example\Emerald 游戏窗口.frm", VBIDEPath & "Template\Forms\Emerald 游戏窗口.frm"
    
    Call FakeSleep
    
    SetupPage.SetupInfo = "正在复制：Visual Basic 6 模板文件（2/2）"
    SetupPage.Progress = 0.9
    FileCopy App.path & "\Example\Emerald 页面.cls", VBIDEPath & "Template\Classes\Emerald 页面.cls"
    
    Call FakeSleep
    
    SetupPage.SetupInfo = "收尾"
    SetupPage.Progress = 1
    
    SetupErr = Err.Number
End Sub
Sub CheckVersion()
    On Error Resume Next
    Dim exeP As String, sh As String
    exeP = """" & App.path & "\Builder.exe" & """"
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    sh = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\version")
    
    If sh <> "" Then
        If Val(sh) <> Version Then
            If Dialog("更新可用", "使用前需要更新你的Emerald。", "更新", "稍后") <> 1 Then Unload MainWindow: End
            Call Setup
            Dialog "更新", "更新成功，请重新启动本程序。", "好的"
            Unload MainWindow: End
        End If
    End If
End Sub
Sub Repair()
    If InstalledPath = "" Then Exit Sub
    
    If Dir(InstalledPath) = "" Then
        ECore.NewTransform transFadeIn, 700, "WelcomePage": Repaired = True
    End If
End Sub
Sub Main2()
    Dim targetEXE As String
    targetEXE = App.path & "\" & App.EXEName & ".exe"
    'targetEXE = "C:\Program Files\Minesweeper\Uninstall.exe"
    'targetEXE = "D:\MyDoc\Emerald\Export\Minesweeper - 安装包.exe"
    
    PackPos = FindPackage(targetEXE, 598000)
    
    MainWindow.Show
    ECore.Display
    DoEvents
    
    LnkSwitch = True
    
    If LCase(Trim(Replace(Command$, """", ""))) = "-uninstallgame" Then
UninstallGame:
        If Dialog("卸载", "确实要卸载该游戏吗？", "是", "手滑") <> 1 Then Unload MainWindow: End
        CmdMark = "Uninstall"
        SetupMode = True
        ECore.NewTransform , 700, "SetupPage"
        Call UninPack
        Exit Sub
    End If
    
    If PackPos <> -1 Then
        '从指定位置把安装包分离出来
        Dim tempPath As String, Data() As Byte, data2() As Byte
        tempPath = VBA.Environ("temp")
        If Dir(tempPath & "\setuppack.emrpack") <> "" Then Kill tempPath & "\setuppack.emrpack"
        ReDim Data(FileLen(targetEXE) - 1)
        ReDim data2(UBound(Data) - PackPos)
        Open targetEXE For Binary As #1
        Get #1, , Data
        Close #1
        CopyMemory data2(0), Data(PackPos), UBound(Data) - PackPos + 1
        ReDim Preserve Data(PackPos - 1)
        Open tempPath & "\setuppack.emrpack" For Binary As #1
        Put #1, , data2
        Close #1
        Open tempPath & "\emrtempUninstall.exe" For Binary As #1
        Put #1, , Data
        Close #1
        Open tempPath & "\setuppack.emrpack" For Binary As #1
        Get #1, , SPackage
        Close #1
        If UBound(SPackage.Files) = 1 Then
            If SPackage.Files(1).path = "setup.config" Then
                '执行卸载程序
                Open App.path & "\setup.config" For Binary As #1
                Put #1, , SPackage.Files(1).Data
                Close #1
                GoTo UninstallGame
            End If
        End If
        
        If SPackage.Files(0).path <> "" Then
            Open tempPath & "\setupappicon.png" For Binary As #1
            Put #1, , SPackage.Files(0).Data
            Close #1
            WelcomePage.Page.Res.newImage tempPath & "\setupappicon.png", 128, 128
        End If
        SetupMode = True
        SSetupPath = "C:\Program Files\" & SPackage.GameName
        Kill tempPath & "\setuppack.emrpack"
        ECore.NewTransform , 700, "WelcomePage"
        Exit Sub
    End If
    
    Call CheckUpdate2
    Call GetVBIDEPath
    Call GetInstalledPath
    Call Repair
    
    If Repaired Then Exit Sub
    
    Cmd = Replace(Command$, """", "")
    'Cmd = "E:\Error 404\Muing III"
    'Cmd = "E:\Error 404\魔兽混战3"
    'Cmd = "pC:\Users\Error404\Desktop\Project\魔兽混战3"
    Dim pmode As Boolean
    If Left(Cmd, 1) = "p" Then pmode = True: Cmd = Right(Cmd, Len(Cmd) - 1)
    
    If Cmd <> "" Then
        Dim appn As String, f As String, t As String, p As String
        Dim nList As String, xinfo As String, info() As String
        p = Cmd

        If p = "-uninstall" Then
            ECore.NewTransform transFadeIn, 700, "WelcomePage"
            Exit Sub
        End If
        
        Call CheckVersion
        
        If Dir(p & "\.emerald") <> "" Then
            Open p & "\.emerald" For Input As #1
            Do While Not EOF(1)
            Line Input #1, t
            xinfo = xinfo & t & vbCrLf
            Loop
            Close #1
            If Dir(p & "\core", vbDirectory) = "" Then MkDir p & "\core"
            If Dir(p & "\.emr", vbDirectory) = "" Then MkDir p & "\.emr"
            If Dir(p & "\.emr\backup", vbDirectory) = "" Then MkDir p & "\.emr\backup"
            If Dir(p & "\.emr\cache", vbDirectory) = "" Then MkDir p & "\.emr\cache"
            If Dir(p & "\assets\debug", vbDirectory) = "" Then MkDir p & "\assets\debug"
            If Dir(p & "\animation", vbDirectory) = "" Then MkDir p & "\animation"
            If Dir(p & "\music", vbDirectory) = "" Then MkDir p & "\music"
            info = Split(xinfo, vbCrLf)
        End If
        
        If Dir(p & "\core\GCore.bas") <> "" Then
            Dim sw2 As String
            If UBound(info) >= 2 Then sw2 = Trim(info(2))
            If Val(info(0)) < Version Or sw2 = "True" Then
                If pmode Then
                    Dialog "请更新", "请保证您的工程正在使用最新版Emerald。", "OK"
                    Unload MainWindow: End
                End If
                ECore.NewTransform , 700, "UpdatePage"
                AppInfo = info
                UpdatePage.GetWarnStr
                Exit Sub
            Else
                If pmode Then
                    If Dialog("打包", "现在开始打包吗？", "好", "不要") <> 1 Then Unload MainWindow: End
                    If Dir(Cmd & "\app.exe") = "" Then
                        Dialog "警告", "找不到游戏主程序：app.exe，请设置。", "行"
                        Unload MainWindow: End
                    End If
                    Dim QQ As Long, Maker As String, Name As String, Describe As String, GVersion As String
                    Dim tempr As String
                    Open Cmd & "\" & Dir(Cmd & "\*.vbp") For Input As #1
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
                    If Name = "" Then
                        Dialog "警告", "游戏名称不能为空。", "行"
                        Unload MainWindow: End
                    End If
                    MakePackage Cmd, Maker, Name, GVersion, Describe, QQ
                    CreateFolder GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\"
                    If Dir(GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & Name & " - 安装包.exe") <> "" Then Kill GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & Name & " - 安装包.exe"
                    Open VBA.Environ("temp") & "\copyemr.cmd" For Output As #1
                    Print #1, "@echo off"
                    Print #1, "echo Emerald Package Toolkit , Version: " & Version
                    Print #1, "echo Building Installer..."
                    Print #1, "ping localhost -n 3 > nul"
                    Print #1, "copy """ & targetEXE & """ /b + """ & VBA.Environ("temp") & "\emrpack"" /b """ & GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & Name & " - 安装包.exe"""
                    Close #1
                    ShellExecuteA 0, "open", VBA.Environ("temp") & "\copyemr.cmd", "", "", SW_SHOW
                    Do While Dir(GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & Name & " - 安装包.exe") = ""
                        Sleep 10: DoEvents
                        ECore.Display
                    Loop
                    Dialog "恭喜", "安装包制作成功", "好的"
                    ShellExecuteA 0, "open", "explorer.exe", "/select,""" & GetSpecialDir(MYDOCUMENTS) & "\Emerald\Export\" & Name & " - 安装包.exe" & """", "", SW_SHOW
                    Unload MainWindow: End
                End If
                Dialog "无操作", "你的工程已经在使用最新的Emerald了。", "手滑"
                Unload MainWindow: End
            End If
        End If

        appn = InputAsk("创建工程", "输入你的可爱的工程名称(*^▽^*)~", "完成", "取消")
        If CheckFileName(appn) = False Or appn = "" Then Dialog "愤怒", "错误的工程名称。", "诶？": Unload MainWindow: End
        
        Open App.path & "\Example\example.vbp" For Input As #1
        Do While Not EOF(1)
        Line Input #1, t
        f = f & t & vbCrLf
        Loop
        Close #1
        
        f = Replace(f, "{app}", appn)
        
        Open p & "\" & appn & ".vbp" For Output As #1
        Print #1, f
        Close #1
        '先下手忽略Emerald文件夹
        Open p & "\.gitignore" For Output As #1
        Print #1, ".emr/*"
        Close #1
        
SkipName:
        If Dir(p & "\core", vbDirectory) = "" Then MkDir p & "\core"
        If Dir(p & "\.emr", vbDirectory) = "" Then MkDir p & "\.emr"
        If Dir(p & "\.emr\backup", vbDirectory) = "" Then MkDir p & "\.emr\backup"
        If Dir(p & "\.emr\cache", vbDirectory) = "" Then MkDir p & "\.emr\cache"
        If Dir(p & "\assets", vbDirectory) = "" Then MkDir p & "\assets"
        If Dir(p & "\assets\debug", vbDirectory) = "" Then MkDir p & "\assets\debug"
        If Dir(p & "\music", vbDirectory) = "" Then MkDir p & "\music"
        
        CopyInto App.path & "\core", p & "\core", True
        CopyInto App.path & "\assets\debug", p & "\assets\debug"
        CopyInto App.path & "\framework", p
        
        Open p & "\.emerald" For Output As #1
        Print #1, Version 'version
        Print #1, Now 'Update Time
        Print #1, False
        Close #1
        
    Else
        
        If InstalledPath <> "" Then
            If (Not IsUpdate) Then
                ECore.NewTransform transFadeIn, 700, "WelcomePage": Exit Sub
            Else
                ECore.NewTransform transFadeIn, 700, "WelcomePage": Exit Sub
            End If
        End If
        
        If InstalledPath = "" Then
            ECore.NewTransform transFadeIn, 700, "WelcomePage": Exit Sub
        End If
        
    End If
    
    Unload MainWindow: End
End Sub
Function InputAsk(t As String, C As String, ParamArray B()) As String
    InputAsk = InputBox(C, t)
End Function
Function Dialog(t As String, C As String, ParamArray B()) As Integer
    Dim b2(), last As String
    b2 = B
    
    last = ECore.ActivePage
    DialogPage.NewDialog t, C, b2
    
    Do While DialogPage.Key = 0
        ECore.Display
        Sleep 10: DoEvents
    Loop
    
    Dialog = DialogPage.Key
    ECore.NewTransform transFadeIn, 700, last
End Function
Function CompareFolder(Src As String, Dst As String) As String
    Dim f As String, fs() As String
    f = Dir(Src & "\")
    
    ReDim fs(0)
    
    Do While f <> ""
        ReDim Preserve fs(UBound(fs) + 1)
        fs(UBound(fs)) = f
        f = Dir()
    Loop
    
    For I = 1 To UBound(fs)
        If Dir(Dst & "\" & fs(I)) = "" Then
            CompareFolder = CompareFolder & fs(I) & vbCrLf
        End If
    Next
End Function

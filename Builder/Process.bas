Attribute VB_Name = "Process"
'Emerald 相关代码

Public VBIDEPath As String, InstalledPath As String, IsUpdate As Boolean
Public WelcomePage As New WelcomePage, TitleBar As New TitleBar, SetupPage As SetupPage, WaitPage As WaitPage, DialogPage As DialogPage, UpdatePage As UpdatePage
Public Tasks() As String
Public NewVersion As Long
Public CmdMark As String, SetupErr As Long, Repaired As Boolean
Public AppInfo() As String
Public Cmd As String
Public Sub CheckUpdate()
    On Error GoTo ErrHandle
    
    Dim WSHShell As Object, temp As String
    Set WSHShell = CreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\version")
    IsUpdate = (Val(temp) <> Version)
    
ErrHandle:
    
End Sub
Public Sub GetInstalledPath()
    On Error GoTo ErrHandle
    
    Dim WSHShell As Object, temp As String
    Set WSHShell = CreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\icon")
    InstalledPath = Replace(temp, """", "")
    
ErrHandle:
    
End Sub
Public Sub GetVBIDEPath()
    On Error GoTo ErrHandle
    
    Dim WSHShell As Object, temp As String, temp2() As String
    Set WSHShell = CreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\VisualBasic.Project\shell\open\command\")
    temp2 = Split(temp, "vb6.exe")
    VBIDEPath = Replace(temp2(0), """", "")
    
ErrHandle:
    If Err.Number <> 0 Then
        Dialog "迷路", "获取VB6路径失败，请确认您的电脑上已经安装VB6（非精简版）。" & vbCrLf & vbCrLf & _
               "注意：Emerald只适用于VB6", "好吧"
    End If
End Sub
Public Function CheckFileName(name As String) As Boolean
    CheckFileName = ((InStr(name, "*") Or InStr(name, "\") Or InStr(name, "/") Or InStr(name, ":") Or InStr(name, "?") Or InStr(name, """") Or InStr(name, "<") Or InStr(name, ">") Or InStr(name, "|") Or InStr(name, " ") Or InStr(name, "!") Or InStr(name, "-") Or InStr(name, "+") Or InStr(name, "#") Or InStr(name, "@") Or InStr(name, "$") Or InStr(name, "^") Or InStr(name, "&") Or InStr(name, "(") Or InStr(name, ")")) = 0)
    Dim t As String
    If name <> "" Then t = Left(name, 1)
    CheckFileName = CheckFileName And (Trim(Str(Val(t))) <> t)
End Function
Sub Uninstall()
    'If Dialog("卸载", "Emerald Builder 已经安装，你希望删除它吗？", "卸载", "手滑") <> 1 Then End
    On Error Resume Next
    
    SetupPage.SetupInfo = "正在创建：WScript.Shell对象"
    SetupPage.Progress = 0.1
    Call FakeSleep
    
    Set WSHShell = CreateObject("WScript.Shell")
    
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
    For i = 1 To Counts
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
    Set WSHShell = CreateObject("WScript.Shell")

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
    Set WSHShell = CreateObject("WScript.Shell")
    
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
Public Sub CheckOnLineUpdate()
    On Error Resume Next
    
    Call FakeSleep(300)
    
    If InternetGetConnectedState(0&, 0&) = 0 Then
        NewVersion = 3
        Exit Sub
    End If
    
    Dim data As New GSaving
    data.Create "Emerald.Core", "Emerald.Core"
    If Now - CDate(data.GetData("UpdateTime")) >= UpdateCheckInterval Or data.GetData("UpdateAble") = 1 Then
        data.PutData "UpdateTime", Now
        
        Dim xmlHttp As Object, Ret As String, Start As Long
        Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        xmlHttp.Open "GET", "https://raw.githubusercontent.com/Red-Error404/Emerald/master/Version.txt", True
        xmlHttp.send
        
        Start = GetTickCount
        Do While xmlHttp.ReadyState <> 4
            If GetTickCount - Start >= UpdateTimeOut Then
                NewVersion = 3
                Exit Sub
            End If
            ECore.Display
            Sleep 10: DoEvents
        Loop
        Ret = xmlHttp.responseText
        Set xmlHttp = Nothing

        NewVersion = Val(Ret)
        data.PutData "UpdateAble", 1

    Else
    
        NewVersion = Version
        
    End If
End Sub
Sub Main()
    MainWindow.Show
    
    Call CheckUpdate
    Call GetVBIDEPath
    Call GetInstalledPath
    Call Repair
    
    If Repaired Then Exit Sub
    
    Cmd = Replace(Command$, """", "")
    Cmd = "E:\Error 404\魔兽混战3"
    
    If Cmd <> "" Then
        Dim appn As String, f As String, t As String, p As String
        Dim nList As String, xinfo As String, info() As String
        p = Cmd

        If p = "-uninstall" Then Call Uninstall
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
            If Dir(p & "\music", vbDirectory) = "" Then MkDir p & "\music"
            info = Split(xinfo, vbCrLf)
        End If
        
        If Dir(p & "\core\GCore.bas") <> "" Then
            Dim sw2 As Boolean
            If UBound(info) >= 2 Then sw2 = Val(info(2))
            If Val(info(0)) < Version Or sw2 Then
                ECore.NewTransform , 700, "UpdatePage"
                AppInfo = info
                Exit Sub
            Else
                Dialog "无操作", "你的工程已经在使用最新的Emerald了。", "手滑"
                Unload MainWindow: End
            End If
        End If

        appn = InputAsk("创建工程", "输入你的可爱的工程名称(*^▽^*)~", "完成", "取消")
        If CheckFileName(appn) = False Or appn = "" Then Dialog "愤怒", "错误的工程名称。", "诶？": Unload MainWindow: End
        
        Open App.path & "\example.vbp" For Input As #1
        Do While Not EOF(1)
        Line Input #1, t
        f = f & t & vbCrLf
        Loop
        Close #1
        
        f = Replace(f, "{app}", appn)
        
        Open p & "\" & appn & ".vbp" For Output As #1
        Print #1, f
        Close #1
            
SkipName:
        If Dir(p & "\core", vbDirectory) = "" Then MkDir p & "\core"
        If Dir(p & "\.emr", vbDirectory) = "" Then MkDir p & "\.emr"
        If Dir(p & "\.emr\backup", vbDirectory) = "" Then MkDir p & "\.emr\backup"
        If Dir(p & "\.emr\cache", vbDirectory) = "" Then MkDir p & "\.emr\cache"
        If Dir(p & "\assets\debug", vbDirectory) = "" Then MkDir p & "\assets\debug"
        If Dir(p & "\music", vbDirectory) = "" Then MkDir p & "\music"
        
        CopyInto App.path & "\core", p & "\core"
        CopyInto App.path & "\assets\debug", p & "\assets\debug"
        CopyInto App.path & "\framework", p
        
        Open p & "\.emerald" For Output As #1
        Print #1, Version 'version
        Print #1, Now 'Update Time
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
Function InputAsk(t As String, c As String, ParamArray b()) As String
    InputAsk = InputBox(c, t)
End Function
Function Dialog(t As String, c As String, ParamArray b()) As Integer
    Dim b2(), last As String
    b2 = b
    
    last = ECore.ActivePage
    DialogPage.NewDialog t, c, b2
    
    Do While DialogPage.Key = 0
        ECore.Display
        Sleep 10: DoEvents
    Loop
    
    Dialog = DialogPage.Key
    ECore.NewTransform transFadeIn, 700, last
End Function
Sub CopyInto(Src As String, Dst As String)
    Dim f As String, p As Boolean
    p = Dir(Dst & "\Core.bas") <> ""
    f = Dir(Src & "\")
    Do While f <> ""
        If f = "Core.bas" Then
            If p Then GoTo skip
        End If
        FileCopy Src & "\" & f, Dst & "\" & f
        DoEvents
skip:
        f = Dir()
    Loop
End Sub
Function CompareFolder(Src As String, Dst As String) As String
    Dim f As String, fs() As String
    f = Dir(Src & "\")
    
    ReDim fs(0)
    
    Do While f <> ""
        ReDim Preserve fs(UBound(fs) + 1)
        fs(UBound(fs)) = f
        f = Dir()
    Loop
    
    For i = 1 To UBound(fs)
        If Dir(Dst & "\" & fs(i)) = "" Then
            CompareFolder = CompareFolder & fs(i) & vbCrLf
        End If
    Next
End Function

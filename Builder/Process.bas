Attribute VB_Name = "Process"
Public Const Version As Long = 19050108
Public VBIDEPath As String
Public Sub GetVBIDEPath()
    On Error GoTo errHandle
    
    Dim WSHShell As Object, temp As String, temp2() As String
    Set WSHShell = CreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\VisualBasic.Project\shell\open\command\")
    temp2 = Split(temp, "vb6.exe")
    VBIDEPath = Replace(temp2(0), """", "")
    
errHandle:
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
    If Dialog("卸载", "Emerald Builder 已经安装，你希望删除它吗？", "卸载", "手滑") <> 1 Then End

    On Error Resume Next
    
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\version"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\"
    
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\"
    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\"
    
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayIcon"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayName"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayVersion"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\Publisher"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\URLInfoAbout"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\UninstallString"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\InstallLocation"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\"
    
    Kill VBIDEPath & "Template\Forms\Emerald 游戏窗口.frm"
    Kill VBIDEPath & "Template\Classes\Emerald 页面.cls"
    
    If Err.Number <> 0 Then
        MsgBox "卸载过程中部分步骤出现错误，可能需要您手动确认删除。", 64, "再见"
    End If
    
    Dialog "再见", "Emerald Builder 已经从你的电脑上删除。", "再见"
    
    End
End Sub
Sub Setup()
    Dim exeP As String
    exeP = """" & App.Path & "\Builder.exe" & """"
    Set WSHShell = CreateObject("WScript.Shell")
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\", "在此处创建/更新Emerald工程"
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon", exeP
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\version", Version
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\", exeP & " ""%v"""
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\", "在此处创建/更新Emerald工程"
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon", exeP
    
    WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\", exeP & " ""%v"""
    
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayIcon", exeP
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayName", "Emerald"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayVersion", "Indev " & Version
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\Publisher", "Error 404"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\InstallLocation", App.Path
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\URLInfoAbout", "http://red-error404.github.io/233"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\UninstallString", exeP & " ""-uninstall"""
    
    FileCopy App.Path & "\Example\Emerald 游戏窗口.frm", VBIDEPath & "Template\Forms\Emerald 游戏窗口.frm"
    FileCopy App.Path & "\Example\Emerald 页面.cls", VBIDEPath & "Template\Classes\Emerald 页面.cls"
    
End Sub
Sub CheckVersion()
    On Error Resume Next
    Dim exeP As String, sh As String
    exeP = """" & App.Path & "\Builder.exe" & """"
    Set WSHShell = CreateObject("WScript.Shell")
    
    sh = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\version")
    
    If sh <> "" Then
        If Val(sh) <> Version Then
            If Dialog("更新可用", "使用前需要更新你的Emerald。", "更新", "稍后") <> 1 Then End
            Call Setup
            Dialog "更新", "更新成功，请重新启动本程序。", "好的"
            End
        End If
    End If
End Sub
Sub Main()
    Call GetVBIDEPath
    
    If Command$ <> "" Then
        Dim appn As String, f As String, t As String, p As String
        Dim nList As String, xinfo As String, info() As String
        p = Replace(Command$, """", "")

        If p = "-uninstall" Then Call Uninstall
        Call CheckVersion
        
        If Dir(p & "\.emerald") <> "" Then
            Open p & "\.emerald" For Input As #1
            Do While Not EOF(1)
            Line Input #1, t
            xinfo = xinfo & t & vbCrLf
            Loop
            Close #1
            info = Split(xinfo, vbCrLf)
        End If
        
        If Dir(p & "\core\GCore.bas") <> "" Then
            If Val(info(0)) < Version Then
                nList = nList & CompareFolder(App.Path & "\core", p & "\core") & vbCrLf
                If nList = vbCrLf Then
                    Dialog "工程更新", "我们已将最新的文件复制到你的文件夹中，本次没有新增的文件。", "好的"
                Else
                    Dialog "工程更新", "你的工程已经创建，我们已将最新的文件复制到你的文件夹中，你可以稍后引用它们。" & vbCrLf & vbCrLf & "注意：以下是更新Emerald后新增的文件，需要你手动引用（位于目录下的Core文件夹）：" & vbCrLf & nList, "收到！"
                End If
                GoTo SkipName
            Else
                Dialog "无操作", "你的工程已经在使用最新的Emerald了。", "手滑"
                Exit Sub
            End If
        End If

        appn = InputAsk("创建工程", "输入你的可爱的工程名称(*^▽^*)~", "完成", "取消")
        If CheckFileName(appn) = False Or appn = "" Then Dialog "愤怒", "错误的工程名称。", "诶？": Exit Sub
        
        Open App.Path & "\example.vbp" For Input As #1
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
        CopyInto App.Path & "\core", p & "\core"
        If Dir(p & "\assets", vbDirectory) = "" Then MkDir p & "\assets"
        If Dir(p & "\assets\debug", vbDirectory) = "" Then MkDir p & "\assets\debug"
        CopyInto App.Path & "\assets\debug", p & "\assets\debug"
        CopyInto App.Path & "\framework", p
        If Dir(p & "\music", vbDirectory) = "" Then MkDir p & "\music"
        
        Open p & "\.emerald" For Output As #1
        Print #1, Version 'version
        Print #1, Now 'Update Time
        Close #1
        
    Else
        Dim exeP As String, sh As String
        exeP = """" & App.Path & "\Builder.exe" & """"
        Set WSHShell = CreateObject("WScript.Shell")
        
        On Error GoTo FailRead
        sh = WSHShell.RegRead("HKEY_CLASSES_ROOT\Directory\shell\emerald\version")
FailRead:
        
        On Error GoTo FailOper
        
        If sh <> "" Then
            If Val(sh) = Version Then
                Call Uninstall
                End
            Else
                If Dialog("更新可用", "按下确定后更新你的 Emerald Builder .", "确定", "取消") <> 1 Then Exit Sub
            End If
        End If
        
        Call Setup
        
        Dialog "成功", "Emerald Builder 成功安装在你的电脑上。", "好"
        
FailOper:
        If Err.Number <> 0 Then Dialog "错误", "出了一些意外，无法完成部分操作。" & vbCrLf & Err.Description & "(" & Err.Number & ")", "好吧"
    End If
End Sub
Function InputAsk(t As String, c As String, ParamArray b()) As String
    Dim w As New MainWindow, b2()
    b2 = b
    
    w.NewDialog t, c, "", True, b2
    w.Show
    
    Do While w.Visible
        DoEvents
    Loop
    
    If w.Key = 1 Then InputAsk = w.InputBox.Content
    Unload w
End Function
Function Dialog(t As String, c As String, ParamArray b()) As Integer
    Dim w As New MainWindow, b2()
    b2 = b
    
    w.NewDialog t, c, "", False, b2
    w.Show
    
    Do While w.Visible
        DoEvents
    Loop
    
    Dialog = w.Key
    Unload w
End Function
Sub CopyInto(Src As String, Dst As String)
    Dim f As String
    f = Dir(Src & "\")
    Do While f <> ""
        FileCopy Src & "\" & f, Dst & "\" & f
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

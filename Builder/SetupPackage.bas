Attribute VB_Name = "SetupPackage"
Public Type EFile
    path As String
    data() As Byte
End Type
Public Type EPackage
    AHead(10) As Byte
    GameName As String
    GameVersion As String
    GameDescribe As String
    MakerQQ As Long
    Maker As String
    files() As EFile
End Type
Public SPackage As EPackage, SetupMode As Boolean
Public Function FindPackage(ByVal File As String, Start As Long) As Long
    Dim Package As EPackage, pos As Long
    With Package                    '设置文件头
        .AHead(0) = 40
        .AHead(1) = 4
        .AHead(2) = 233
        '====================================
        '   文件格式标识
            .AHead(3) = 19
            .AHead(4) = 6
            .AHead(5) = 1
        '====================================
        .AHead(6) = 136
        .AHead(7) = 177
        .AHead(8) = 82
        .AHead(9) = 19
        .AHead(10) = 233
    End With
    
    Dim data As Byte, i As Long
    FindPackage = -1
    Open File For Binary As #1
    For i = Start To FileLen(File)
        Get #1, i, data
        If data = Package.AHead(pos) Then
            pos = pos + 1
            If pos = 11 Then Exit For
        Else
            pos = 0
        End If
    Next
    Close #1
    
    If pos = 11 Then FindPackage = i - 11
End Function
Public Sub MakePackage(ByVal path As String, GMaker As String, GName As String, GVersion As String, GDescribe As String, QQ As Long)
    If Right(path, 1) <> "\" Then path = path & "\"
    
    Dim files() As String
    files = DirAllFiles(path)
    
    Dim Package As EPackage
    With Package                    '设置文件头
        .AHead(0) = 40
        .AHead(1) = 4
        .AHead(2) = 233
        '====================================
        '   文件格式标识
            .AHead(3) = 19
            .AHead(4) = 6
            .AHead(5) = 1
        '====================================
        .AHead(6) = 136
        .AHead(7) = 177
        .AHead(8) = 82
        .AHead(9) = 19
        .AHead(10) = 233
    End With
    
    With Package                    '设置开发者信息
        .GameDescribe = GDescribe
        .GameName = GName
        .GameVersion = GVersion
        .MakerQQ = QQ
        .Maker = GMaker
    End With
    
    Dim data() As Byte
    
    For i = 1 To UBound(files)      '替换为相对路径
        files(i) = Right(files(i), Len(files(i)) - Len(path))
    Next
    
    ReDim Package.files(0)
    
    For i = 1 To UBound(files)
        If LCase(files(i)) = "app.png" Then
            ReDim data(FileLen(path & "app.png") - 1)
            Open path & "app.png" For Binary As #1
            Get #1, , data
            Close #1
            With Package
                .files(0).data = data
                .files(0).path = "app.png"
            End With
            Exit For
        End If
    Next
    
    For i = 1 To UBound(files)
        '排除Visual Basic6代码和Emerald设置文件
        If Not ((LCase(files(i)) Like "*.vbp") Or (LCase(files(i)) Like "*.vbw") Or (LCase(files(i)) Like "*.vbg") Or _
                (LCase(files(i)) Like "*.bas") Or _
                (LCase(files(i)) Like "*.frm") Or (LCase(files(i)) Like "*.frx") Or _
                (LCase(files(i)) Like "*.cls") Or _
                (LCase(files(i)) = ".emerald")) Then
            ReDim data(FileLen(path & files(i)) - 1)
            Open path & files(i) For Binary As #1
            Get #1, , data
            Close #1
            With Package
                ReDim Preserve .files(UBound(.files) + 1)
                .files(UBound(.files)).data = data
                .files(UBound(.files)).path = files(i)
            End With
            Call FakeSleep(1)
        End If
    Next
    
    '导出.emrpack文件
    If Dir(VBA.Environ("temp") & "\emrpack") <> "" Then Kill VBA.Environ("temp") & "\emrpack"
    Open VBA.Environ("temp") & "\emrpack" For Binary As #1
    Put #1, , Package
    Close #1
End Sub
Function DirAllFiles(ByVal path As String) As String()
    Dim DirTasks() As String, File As String, Folder As String
    Dim FileList() As String
    ReDim DirTasks(1), FileList(0)
    If Right(path, 1) <> "\" Then path = path & "\"
    DirTasks(1) = path
    Do While UBound(DirTasks) > 0
        File = Dir(DirTasks(1))
        Do While File <> ""
            ReDim Preserve FileList(UBound(FileList) + 1)
            FileList(UBound(FileList)) = DirTasks(1) & File
            File = Dir()
            DoEvents
        Loop
        Folder = Dir(DirTasks(1), vbDirectory)
        Do While Folder <> ""
            If Folder <> "." And Folder <> ".." And (Not Folder Like "*.*") Then
                ReDim Preserve DirTasks(UBound(DirTasks) + 1)
                DirTasks(UBound(DirTasks)) = DirTasks(1) & Folder & "\"
            End If
            Folder = Dir(, vbDirectory)
            DoEvents
        Loop
        DirTasks(1) = DirTasks(UBound(DirTasks))
        ReDim Preserve DirTasks(UBound(DirTasks) - 1)
    Loop
    DirAllFiles = FileList
End Function
Sub CreateFolder(ByVal path As String)
    Dim temp() As String, NowPath As String
    If Right(path, 1) <> "\" Then path = path & "\"
    temp = Split(path, "\")
    For i = 0 To UBound(temp) - 1
        If temp(i) Like "*.*" Then Exit Sub
        NowPath = NowPath & temp(i) & "\"
        If Dir(NowPath, vbDirectory) = "" Then MkDir NowPath
    Next
End Sub
Sub UninPack()
    On Error Resume Next

    Dim path As String, te As String
    path = App.path
    
    Set WSHShell = CreateObject("WScript.Shell")
    
    Open path & "\setup.config" For Input As #1
    Line Input #1, te
    SPackage.GameName = te
    Close #1
    
    SetupPage.SetupInfo = "正在删除：软件信息"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayIcon"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayName"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayVersion"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\Publisher"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\InstallLocation"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\URLInfoAbout"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\UninstallString"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\"
    
    Call FakeSleep
    
    Dim files() As String
    files = DirAllFiles(App.path)
    
    For i = 1 To UBound(files)
        SetupPage.SetupInfo = "正在删除：" & Replace(files(i), App.path & "\", "")
        If files(i) <> "Uninstall.exe" Then Kill files(i)
        SetupPage.Progress = i / UBound(files)
        Call FakeSleep(1)
    Next
        
    SetupErr = Err.Number
End Sub
Sub SetupPack()
    On Error Resume Next

    Dim path As String
    path = "C:\Program Files\" & SPackage.GameName & "\"
    
    Set WSHShell = CreateObject("WScript.Shell")
    
    SetupPage.SetupInfo = "正在注册：软件信息"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayIcon", """" & path & "App.exe" & """"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayName", SPackage.GameName
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayVersion", SPackage.GameVersion
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\Publisher", "QQ " & SPackage.MakerQQ
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\InstallLocation", path
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\URLInfoAbout", ""
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\UninstallString", """" & path & "Uninstall.exe" & """ ""-uninstallgame"""
    
    Call FakeSleep
    
    CreateFolder path
    
    For i = 1 To UBound(SPackage.files)
        SetupPage.SetupInfo = "正在写入：" & SPackage.files(i).path
        CreateFolder path & SPackage.files(i).path
        Open path & SPackage.files(i).path For Binary As #1
        Put #1, , SPackage.files(i).data
        Close #1
        SetupPage.Progress = i / UBound(SPackage.files)
        Call FakeSleep(1)
    Next
    
    FileCopy VBA.Environ("temp") & "\emrtempUninstall.exe", path & "Uninstall.exe"
    
    Open path & "setup.config" For Output As #1
    Print #1, SPackage.GameName
    Close #1
    
    On Error Resume Next
    Dim objShell As Object, objShortcut As Object, strStart As String
    Set objShell = CreateObject("WScript.Shell")
    strStart = objShell.SpecialFolders("Desktop") & "\"
    If Dir(strStart & "\" & SPackage.GameName & ".lnk") <> "" Then Exit Sub
    Set objShortcut = objShell.CreateShortcut(strStart & "\" & SPackage.GameName & ".lnk")
    objShortcut.TargetPath = path & "app.exe"
    objShortcut.Arguments = ""
    objShortcut.WindowStyle = 1
    objShortcut.Hotkey = ""
    objShortcut.IconLocation = path & "app.exe"
    objShortcut.Description = SPackage.GameDescribe
    objShortcut.WorkingDirectory = path
    objShortcut.Save
    Set objShell = Nothing
    Set objShortcut = Nothing
    
    SetupPage.SetupInfo = "正在创建：桌面快捷方式"
    Call FakeSleep(100)
    
    SetupErr = Err.Number
    
End Sub

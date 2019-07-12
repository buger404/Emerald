Attribute VB_Name = "SetupPackage"
Public Type EFile
    Path As String
    Data() As Byte
End Type
Public Type EPackage
    AHead(10) As Byte
    GameName As String
    GameVersion As String
    GameDescribe As String
    MakerQQ As Long
    Maker As String
    Files() As EFile
End Type
Public SPackage As EPackage, SetupMode As Boolean
Public SSetupPath As String
Public LogPath As String
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
    
    Dim Data As Byte, I As Long
    FindPackage = -1
    Open File For Binary As #1
    For I = Start To FileLen(File)
        Get #1, I, Data
        If Data = Package.AHead(pos) Then
            pos = pos + 1
            If pos = 11 Then Exit For
        Else
            pos = 0
        End If
    Next
    Close #1
    
    If pos = 11 Then FindPackage = I - 11
End Function
Public Sub MakePackage(ByVal Path As String, GMaker As String, GName As String, GVersion As String, GDescribe As String, QQ As Long)
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    Dim Files() As String
    Files = DirAllFiles(Path)
    
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
    
    Dim Data() As Byte
    
    For I = 1 To UBound(Files)      '替换为相对路径
        Files(I) = Right(Files(I), Len(Files(I)) - Len(Path))
    Next
    
    ReDim Package.Files(0)
    
    For I = 1 To UBound(Files)
        If LCase(Files(I)) = "app.png" Then
            ReDim Data(FileLen(Path & "app.png") - 1)
            Open Path & "app.png" For Binary As #1
            Get #1, , Data
            Close #1
            With Package
                .Files(0).Data = Data
                .Files(0).Path = "app.png"
            End With
            Exit For
        End If
    Next
    
    For I = 1 To UBound(Files)
        '排除Visual Basic6代码和Emerald设置文件
        If Not ((LCase(Files(I)) Like "*.vbp") Or (LCase(Files(I)) Like "*.vbw") Or (LCase(Files(I)) Like "*.vbg") Or _
                (LCase(Files(I)) Like "*.bas") Or _
                (LCase(Files(I)) Like "*.frm") Or (LCase(Files(I)) Like "*.frx") Or _
                (LCase(Files(I)) Like "*.cls") Or _
                (LCase(Files(I)) = ".emerald")) Then
            ReDim Data(FileLen(Path & Files(I)) - 1)
            Open Path & Files(I) For Binary As #1
            Get #1, , Data
            Close #1
            With Package
                ReDim Preserve .Files(UBound(.Files) + 1)
                .Files(UBound(.Files)).Data = Data
                .Files(UBound(.Files)).Path = Files(I)
            End With
            If PackPos = -1 Then WelcomePage.PackText = "Packing '" & Files(I) & "' ..."
        Else
            If PackPos = -1 Then WelcomePage.PackText = "Removing '" & Files(I) & "' ..."
        End If
        ECore.Display: DoEvents
    Next
    
    '导出.emrpack文件
    If PackPos = -1 Then WelcomePage.PackText = "Exporting emrpack ..."
    If Dir(VBA.Environ("temp") & "\emrpack") <> "" Then Kill VBA.Environ("temp") & "\emrpack"
    Open VBA.Environ("temp") & "\emrpack" For Binary As #1
    Put #1, , Package
    Close #1
End Sub
Function DirAllFiles(ByVal Path As String) As String()
    Dim DirTasks() As String, File As String, Folder As String
    Dim FileList() As String
    ReDim DirTasks(1), FileList(0)
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    DirTasks(1) = Path
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
Sub CreateFolder(ByVal Path As String)
    Dim temp() As String, NowPath As String
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    temp = Split(Path, "\")
    For I = 0 To UBound(temp) - 1
        If temp(I) Like "*.*" Then Exit Sub
        NowPath = NowPath & temp(I) & "\"
        If Dir(NowPath, vbDirectory) = "" Then MkDir NowPath
    Next
End Sub
Public Function UninPack() As String
    On Error Resume Next

    Dim Path As String, te As String
    Path = App.Path
    
    Randomize
    LogPath = VBA.Environ("temp") & "\Emerald_Setup_" & Int(Rnd * 999999999 + 1111111111) & ".txt"
    
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    Open Path & "\setup.config" For Input As #1
    Line Input #1, te
    SPackage.GameName = te
    Close #1
    
    Open LogPath For Output As #2
    Print #2, "Emerald Uninstaller Report"
    Print #2, "Game name：" & SPackage.GameName
    Print #2, ""
    ECore.Display: DoEvents
    
    SetupPage.SetupInfo = "Deleting software informations ..."
    Print #2, Now & "    " & "RegDelete: HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayIcon"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayName"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayVersion"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\Publisher"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\InstallLocation"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\URLInfoAbout"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\UninstallString"
    WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\"
    
    Dim Files() As String
    Files = DirAllFiles(App.Path)
    
    For I = 1 To UBound(Files)
        SetupPage.SetupInfo = "Deleting '" & Replace(Files(I), App.Path & "\", "") & "' ..."
        Print #2, Now & "    " & "Delete: " & Files(I)
        If Files(I) <> "Uninstall.exe" Then Kill Files(I)
        SetupPage.Progress = I / UBound(Files)
        ECore.Display: DoEvents
    Next
    
    Close #2
    
    UninPack = Err.Description
End Function
Public Function SetupPack() As String
    On Error Resume Next

    Dim Path As String
    Path = SSetupPath & "\"
    
    Randomize
    LogPath = VBA.Environ("temp") & "\Emerald_Setup_" & Int(Rnd * 999999999 + 1111111111) & ".txt"
    Open LogPath For Output As #2
    Print #2, "Emerald Installer Report"
    Print #2, "Game name : " & SPackage.GameName
    Print #2, ""
    
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    ECore.Display: DoEvents
    
    SetupPage.SetupInfo = "Writting software informations ..."
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayIcon", """" & Path & "App.exe" & """"
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayName", SPackage.GameName
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\DisplayVersion", SPackage.GameVersion
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\Publisher", SPackage.Maker
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\InstallLocation", Path
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\URLInfoAbout", ""
    WSHShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName & "\UninstallString", """" & Path & "Uninstall.exe" & """"
    Print #2, Now & "    " & "RegWrite: HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\" & SPackage.GameName
    
    CreateFolder Path
    Print #2, Now & "    " & "Create: " & Path
    
    For I = 1 To UBound(SPackage.Files)
        SetupPage.SetupInfo = "Writting '" & SPackage.Files(I).Path & "' ..."
        Print #2, Now & "    " & "Write: " & Path & SPackage.Files(I).Path
        CreateFolder Path & SPackage.Files(I).Path
        Open Path & SPackage.Files(I).Path For Binary As #1
        Put #1, , SPackage.Files(I).Data
        Close #1
        SetupPage.Progress = I / UBound(SPackage.Files)
        ECore.Display: DoEvents
    Next
    
    Print #2, Now & "    " & "Copy: " & VBA.Environ("temp") & "\emrtempUninstall.exe" & " -> " & Path & "Uninstall.exe"
    FileCopy VBA.Environ("temp") & "\emrtempUninstall.exe", Path & "Uninstall.exe"
    
    Dim RandomFolder As String
    Randomize
    RandomFolder = Hex(Int(Rnd * 999999999 + 100000000))
    MkDir Path & RandomFolder
    Print #2, Now & "    " & "Create: " & Path & RandomFolder
    
    Open Path & RandomFolder & "\setup.config" For Output As #1
    Print #1, SPackage.GameName
    Close #1
    Print #2, Now & "    " & "Write: " & Path & RandomFolder & "\setup.config"
    MakePackage Path & RandomFolder, "none", "none", "none", "none", 0
    Print #2, Now & "    " & "Package: " & Path & RandomFolder
    
    Open VBA.Environ("temp") & "\copyemr.cmd" For Output As #1
    Print #1, "@echo off"
    Print #1, "copy """ & VBA.Environ("temp") & "\emrtempUninstall.exe" & """ /b + """ & VBA.Environ("temp") & "\emrpack"" /b """ & Path & "Uninstall.exe" & """"
    Close #1
    Print #2, Now & "    " & "Command: " & "copy """ & VBA.Environ("temp") & "\emrtempUninstall.exe" & """ /b + """ & VBA.Environ("temp") & "\emrpack"" /b """ & Path & "Uninstall.exe" & """"
    ShellExecuteA 0, "open", VBA.Environ("temp") & "\copyemr.cmd", "", "", SW_SHOW
    Print #2, Now & "    " & "Run: " & VBA.Environ("temp") & "\copyemr.cmd"
    
    On Error Resume Next
    ECore.Display: DoEvents
    If LnkSwitch Then
        Dim objShell As Object, objShortcut As Object, strStart As String
        Set objShell = PoolCreateObject("WScript.Shell")
        strStart = objShell.SpecialFolders("Desktop") & "\"
        If Dir(strStart & "\" & SPackage.GameName & ".lnk") <> "" Then GoTo last
        Set objShortcut = objShell.CreateShortcut(strStart & "\" & SPackage.GameName & ".lnk")
        objShortcut.TargetPath = Path & "app.exe"
        objShortcut.Arguments = ""
        objShortcut.WindowStyle = 1
        objShortcut.Hotkey = ""
        objShortcut.IconLocation = Path & "app.exe"
        objShortcut.Description = SPackage.GameDescribe
        objShortcut.WorkingDirectory = Path
        objShortcut.Save
        Set objShell = Nothing
        Set objShortcut = Nothing
        Print #2, Now & "    " & "Create: " & strStart & "\" & SPackage.GameName & ".lnk"
        SetupPage.SetupInfo = "Creating desktop shortcut ..."
    End If
last:
    Close #2
    
    SetupPack = Err.Description
End Function

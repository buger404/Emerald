Attribute VB_Name = "BuilderCore"
Dim WSHShell As Object
Public VBIDEPath As String, NewVersion As Long
Public OPath As String
Public Type EmrPConfig
    AFileHeader As String
    Name As String
    Maker As String
    Version As Long
    FUpdate As Boolean
    AssetsPath As String
    MusicPath As String
    LastBranch As String
    Reserved(1000) As Byte
End Type
Public Type EmrFile
    Path As String
    Data() As Byte
    MD5Check As String
End Type
Public Type EmrBackup
    AFileHeader As String
    Date As String
    Files() As EmrFile
End Type
Public EmrPC As EmrPConfig
Public Sub CopyInto(Src As String, Dst As String, Optional WriteCache As Boolean = False)
    Dim f As String, p As Boolean
    p = (Dir(Dst & "\Core.bas") <> "")
    f = Dir(Src & "\")
    Do While f <> ""
        If f = "Core.bas" Then
            If p Then GoTo skip
        End If
        FileCopy Src & "\" & f, Dst & "\" & f
skip:
        f = Dir()
    Loop
End Sub
Public Sub Main()
    OPath = Replace(Trim(Command$), """", "")
    'OPath = "E:\Error 404\testEmr\"
    EmrPC.AFileHeader = "Emerald Project Config File"
    Call GetVBIDEPath
    If VBIDEPath = "" Then
        MsgBox "Emerald requires Visual Basic 6.0 .", 16
        End
    End If
    MainWindow.Show
    EmrPC.Maker = ESave.GetData("Maker")
    If OPath <> "" Then WelcomePage.InitProject
End Sub
Public Sub GetVBIDEPath()
    On Error Resume Next
    
    Dim temp As String, temp2() As String
    If WSHShell Is Nothing Then Set WSHShell = CreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\VisualBasic.Project\shell\open\command\")
    temp2 = Split(temp, "vb6.exe")
    VBIDEPath = Replace(temp2(0), """", "")
    
End Sub
Public Function UninstallOld() As String
    On Error Resume Next
    
    If WSHShell Is Nothing Then Set WSHShell = CreateObject("WScript.Shell")
    
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emerald\icon") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emerald\version") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\version"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emerald\command\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emerald\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\"
    
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\"
    
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emeraldp\icon") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\icon"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emeraldp\command\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\command\"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\shell\emeraldp\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emeraldp\"
    
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\icon") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\icon"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\command\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\command\"
    If IsRegCreated("HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\") Then WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emeraldp\"
    
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayIcon") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayIcon"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayName") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayName"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayVersion") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\DisplayVersion"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\Publisher") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\Publisher"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\URLInfoAbout") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\URLInfoAbout"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\UninstallString") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\UninstallString"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\InstallLocation") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\InstallLocation"
    If IsRegCreated("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\") Then WSHShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall\Emerald\"
    
    If Dir(VBIDEPath & "Template\Forms\Emerald 游戏窗口.frm") <> "" Then Kill VBIDEPath & "Template\Forms\Emerald 游戏窗口.frm"
    If Dir(VBIDEPath & "Template\Forms\Emerald 页面.cls") <> "" Then Kill VBIDEPath & "Template\Classes\Emerald 页面.cls"
    
    UninstallOld = Err.Description
    Err.Clear
End Function
Public Function IsRegCreated(Path As String) As Boolean

    If WSHShell Is Nothing Then Set WSHShell = PoolCreateObject("WScript.Shell")
    
    On Error Resume Next
    Dim temp As String
    
    temp = WSHShell.RegRead(Path)
    
    IsRegCreated = (Err.Number = 0)
    Err.Clear
    
End Function
Public Function OperContentMenu(Remove As Boolean) As String
    On Error Resume Next
    
    Set WSHShell = CreateObject("WScript.Shell")
    
    Dim exeP As String
    exeP = """" & App.Path & "\Builder.exe" & """"
    
    Dim Items(2) As String, List(1) As String, Text(2) As String
    List(0) = "": List(1) = "Background\"
    Items(0) = "icon": Items(1) = "command\": Items(2) = ""
    Text(0) = exeP: Text(1) = exeP & " ""%v""": Text(2) = "Launch Emerald Builder Here"
    
    For I = 0 To UBound(List)
        For S = 0 To UBound(Items)
            If Remove Then
                WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\" & List(I) & "shell\emerald\" & Items(S)
            Else
                WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\" & List(I) & "shell\emerald\" & Items(S), Text(S)
            End If
        Next
    Next
    
    OperContentMenu = Err.Description
    
    Set WSHShell = Nothing
End Function
Public Sub CheckOnLineUpdate()
    On Error Resume Next
    
    If InternetGetConnectedState(0&, 0&) = 0 Then
        NewVersion = 3
        Exit Sub
    End If
    
    Dim Data As New GSaving
    Data.Create "Emerald.Core"
    Data.AutoSave = True
    If Now - CDate(Data.GetData("UpdateTime")) >= UpdateCheckInterval Or Data.GetData("UpdateAble") = 1 Then
        Data.PutData "UpdateTime", Now
        
        Dim xmlHttp As Object, Ret As String, Start As Long
        Set xmlHttp = PoolCreateObject("Microsoft.XMLHTTP")
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
        Data.PutData "UpdateAble", 1
    Else
        NewVersion = Version
    End If
End Sub


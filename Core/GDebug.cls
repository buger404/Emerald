VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald Ïà¹Ø´úÂë

Public Page As GPage, GW As Single, GH As Single
Public PageMark As Long

Public InputingText As String, InputAllow As Boolean

Private Enum colors
    HighLightText = -18176
    DefaultText = -855310
    ContentText = -5592406
    ErrText = -109511
    SuccessText = -16724886
End Enum

Dim Color As Long
Dim ConsoleDC As Long, ConsoleGG As Long, Exec As String, Params() As String
Dim FirstDisplay As Boolean
Dim HistoryCmd(20) As String, CurrentCmd As Integer, ChooseCmd As Integer

Public sy As Single, NeedScroll As Boolean
Public CuY As Long, CuX As Long
Public Sub EndInput()
    SetConsoleColor colors.DefaultText
    WriteWord InputingText
    NewLine
End Sub
Public Sub ApplyCmd()
    Call EndInput
    Select Case Exec
        Case ""
            Dim param() As String
            param = Split(InputingText, " ")
            If UBound(param) = 0 Then ReDim Preserve param(1)
            
            If InputingText = "" Or Exec <> "" Then GoTo SkipRecord
            
            HistoryCmd(CurrentCmd) = InputingText
            
            If CurrentCmd < UBound(HistoryCmd) Then
                CurrentCmd = CurrentCmd + 1
            Else
                For I = 1 To UBound(HistoryCmd)
                    HistoryCmd(I - 1) = HistoryCmd(I)
                Next
                HistoryCmd(UBound(HistoryCmd)) = ""
            End If
    
SkipRecord:
            ChooseCmd = CurrentCmd

            InputAllow = False
            If InputingText <> "" Then
                Exec = param(0)
                Params = param
                Call ReserveExec
            Else
                Call AllowCmd
            End If
            InputingText = ""
    End Select
End Sub
Public Sub ExitExec()
    Call EndInput
    Exec = ""
    Call AllowCmd
End Sub
Public Sub ApplyKey(Key As Integer)
    Select Case Exec
        Case "showfps"
            If LCase(Chr(Key)) = "x" Then
                Call ExitExec
            End If
        Case "mousestate"
            If LCase(Chr(Key)) = "x" Then
                Call ExitExec
            End If
        Case "save"
            If Params(1) = "-changekey2" Then
                If LCase(Chr(Key)) = "y" Then
                    Params(1) = "-changekey3"
                    InputAllow = False
                ElseIf LCase(Chr(Key)) = "n" Then
                    Call ExitExec
                End If
            End If
        Case "backup"
            If Params(1) = "-clear2" Then
                If LCase(Chr(Key)) = "y" Then
                    Params(1) = "-clear3"
                    InputAllow = False
                ElseIf LCase(Chr(Key)) = "n" Then
                    Call ExitExec
                End If
            End If
        Case ""
            If Key = vbKeyUp Then
                ChooseCmd = ChooseCmd - 1
                If ChooseCmd < 0 Then ChooseCmd = 0
                InputingText = HistoryCmd(ChooseCmd)
            End If
            If Key = vbKeyDown Then
                If HistoryCmd(ChooseCmd) = "" Then Exit Sub
                ChooseCmd = ChooseCmd + 1
                If ChooseCmd > UBound(HistoryCmd) Then
                    ChooseCmd = UBound(HistoryCmd)
                    If CurrentCmd = UBound(HistoryCmd) Then
                        InputingText = "": Exit Sub
                    End If
                End If
                If HistoryCmd(ChooseCmd) = "" Then
                    InputingText = "": Exit Sub
                End If
                InputingText = HistoryCmd(ChooseCmd)
            End If
    End Select
End Sub
Public Sub ReserveExec()
    Select Case Exec
        Case "showfps"
            Call AllowInput
            Call NewLine
        Case "mousestate"
            Call AllowInput
            Call NewLine
    End Select
End Sub
Private Sub CheckUpdate(force As Boolean)
    On Error Resume Next
    If InternetGetConnectedState(0&, 0&) = 0 And force = False Then
        SetConsoleColor colors.ErrText
        WriteLine "Emerald-Error: connect to the Internet first ."
        Err.Clear
        Exit Sub
    End If
    
    Dim Data As New GSaving
    Data.Create "Emerald.Core"
    Data.AutoSave = True
    If Now - CDate(Data.GetData("UpdateTime")) >= UpdateCheckInterval Or force Or Data.GetData("UpdateAble") = 1 Then
        Data.PutData "UpdateTime", Now
        
        Dim xmlHttp As Object, Ret As String, Start As Long
        
        SetConsoleColor colors.DefaultText
        WriteLine "Emerald: connecting to the server"
        Start = GetTickCount
        Do While GetTickCount - Start <= 2000
            Call BackLine: Call ClearLine
            WriteLine "Emerald: connecting to the server" & IIf(GetTickCount Mod 1000 > 250, ".", "") & IIf(GetTickCount Mod 1000 > 500, ".", "") & IIf(GetTickCount Mod 1000 > 750, ".", "")
            Debuginfo.UpdateTimer_Timer
            Sleep 10: DoEvents
        Loop
        
        Set xmlHttp = PoolCreateObject("Microsoft.XMLHTTP")
        xmlHttp.Open "GET", "https://raw.githubusercontent.com/Red-Error404/Emerald/master/Version.txt", True
        xmlHttp.send
                     
        Start = GetTickCount
        
        Do While xmlHttp.ReadyState <> 4
            If GetTickCount - Start >= UpdateTimeOut Then
                SetConsoleColor colors.ErrText
                WriteLine "Emerald-Error: failed to connect , time out "
                Err.Clear
                Exit Sub
            End If
            Call BackLine: Call ClearLine
            WriteLine "Emerald: connecting to the server" & IIf(GetTickCount Mod 4000 < 1000, ".", "") & IIf(GetTickCount Mod 4000 < 2000, ".", "") & IIf(GetTickCount Mod 4000 < 3000, ".", "")
            Debuginfo.UpdateTimer_Timer
            Sleep 10: DoEvents
        Loop
        Ret = xmlHttp.responseText
        Set xmlHttp = Nothing
        SetConsoleColor colors.SuccessText
        WriteLine "Emerald-Success: the lastest version from the server " & Val(Ret)
        SetConsoleColor colors.ContentText
        WriteLine "Emerald: current version " & Version
        
        If Val(Ret) > Version And App.LogMode = 0 Then
            SetConsoleColor colors.SuccessText
            WriteLine "Emerald: update available , plz visit 'https://github.com/Red-Error404/Emerald/release' to download ."
            Exit Sub
        End If
    Else
        SetConsoleColor colors.HighLightText
        WriteLine "Emerald-Skip: update has been checked at " & CDate(Data.GetData("UpdateTime")) & "."
    End If
    
    Err.Clear
    Set Data = Nothing
End Sub
Public Function GetVBIDEPath() As String
    On Error GoTo ErrHandle
    
    Dim WSHShell As Object, temp As String, temp2() As String
    Set WSHShell = PoolCreateObject("WScript.Shell")
    
    temp = WSHShell.RegRead("HKEY_CLASSES_ROOT\VisualBasic.Project\shell\open\command\")
    temp2 = Split(temp, "vb6.exe")
    GetVBIDEPath = Replace(temp2(0), """", "")
    
ErrHandle:
    GetVBIDEPath = ""
End Function
Public Sub RunExec()
    If Exec = "" Then Exit Sub

    Dim MissingParam As Integer

    Select Case Exec
        Case "poolinfo"
            Select Case Params(1)
                Case "-help"
                    WriteWord "Show pool information ."
                    Call ExitExec
                    Exit Sub
                Case ""
                    Dim poolinfo() As String
                    poolinfo = Split(GetCountStr, vbCrLf)
                    For I = 0 To UBound(poolinfo)
                        WriteLine poolinfo(I)
                    Next
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
        Case "backup"
            SetConsoleColor colors.ErrText
            WriteLine "Note£ºThis function will be removed on 7.20 ."
            SetConsoleColor colors.DefaultText
            Select Case Params(1)
                Case "-help"
                    WriteLine "Mange your backups ."
                    WriteLine "-clear : delete all your backups <Danger> ."
                    WriteLine "-list : list your backups ."
                    WriteLine "-replace [backupname] : replace your source files with the backup ."
                    WriteLine "-replace2 [backupname] [filename] : replace single source file with the backup ."
                    Call ExitExec
                    Exit Sub
                Case "-replace2"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If UBound(Params) < 3 Then MissingParam = 3: GoTo MissingParams
                    Params(1) = "-replace"
                Case "-replace"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    
                    Params(1) = "-clear2"
                    
                    Dim bFile As String, Size1 As Long, Size2 As Long
                    
                    If Dir(App.path & "\.emr\backup\" & Params(2), vbDirectory) = "" Then
                        SetConsoleColor colors.ErrText
                        WriteWord "backup-Error: the backup not found '" & App.path & "\.emr\backup\" & Params(2) & "'"
                        Call ExitExec
                        Exit Sub
                    End If
                    Dim bStep As Long, bfo As String, bsfo As String, targetFile As String, targetGet As Boolean
                    bFile = Dir(App.path & "\.emr\backup\" & Params(2) & "\"): bfo = "core": bsfo = ""
                    If UBound(Params) = 3 Then targetFile = Params(3): WriteLine "Finding"
                    
ReCopy:
                    'On Error Resume Next
                    Do While bFile <> ""
                        If targetFile <> "" Then
                            BackLine
                            ClearLine
                            WriteWord "Finding" & IIf(GetTickCount Mod 4000 < 1000, ".", "") & IIf(GetTickCount Mod 4000 < 2000, ".", "") & IIf(GetTickCount Mod 4000 < 3000, ".", "")
                            NextLine
                        End If
                        If (targetFile <> "" And targetFile = bFile) Or (targetFile = "") Then
                            Size1 = FileLen(App.path & "\.emr\backup\" & Params(2) & "\" & bsfo & "\" & bFile): Size2 = 0
                            Size2 = FileLen(App.path & "\" & bfo & "\" & bFile)
                            FileCopy App.path & "\.emr\backup\" & Params(2) & "\" & bsfo & "\" & bFile, App.path & "\" & bfo & "\" & bFile
                            SetConsoleColor colors.DefaultText
                            WriteWord "Success : " & bFile & "      "
                            SetConsoleColor IIf(Size1 = Size2, colors.DefaultText, IIf(Size1 > Size2, colors.SuccessText, colors.ErrText))
                            WriteWord Size2 & " bytes -> " & Size1 & " bytes"
                            NewLine
                            targetGet = True
                            If targetFile <> "" Then Exit Do
                        End If
                        bFile = Dir()
                        Debuginfo.UpdateTimer_Timer
                        Sleep 10: DoEvents
                    Loop
                    bStep = bStep + 1
                    If bStep = 1 And (Not targetGet) Then bFile = Dir(App.path & "\.emr\backup\" & Params(2) & "\cache\"): bfo = ".emr\cache": bsfo = "cache"
                    If bStep = 2 And (Not targetGet) Then bFile = Dir(App.path & "\.emr\backup\" & Params(2) & "\project\"): bfo = "": bsfo = "project"
                    If bStep <= 2 And (Not targetGet) Then GoTo ReCopy
                    
                    If targetFile = "" Then
                        SetConsoleColor colors.SuccessText
                        WriteWord "backup-Success: succeed in replacing files with the backup (" & Params(2) & ")"
                    Else
                        If targetGet = False Then
                            SetConsoleColor colors.ErrText
                            WriteWord "backup-Error: file not found '" & targetFile & "'"
                        Else
                            SetConsoleColor colors.SuccessText
                            WriteWord "backup-Success: succeed in replacing the file with the backup (" & Params(2) & ")"
                        End If
                    End If
                    
                    Call ExitExec
                    Exit Sub
                Case "-clear"
                    SetConsoleColor colors.ErrText
                    WriteLine "Danger !"
                    WriteLine "The consequences are at one's own expense !!!"
                    WriteLine "All your backups will be erased from your drive forever ."
                    SetConsoleColor colors.DefaultText
                    WriteLine "Target folder : " & App.path & "\.emr\backup\"
                    WriteLine "Erase them any way ? [Y/N]"
                    AllowInput
                    Params(1) = "-clear2"
                Case "-clear2"
                Case "-clear3"
                    Dim FSO3 As Object
                    Set FSO3 = PoolCreateObject("Scripting.FileSystemObject")
                    FSO3.DeleteFolder App.path & "\.emr\backup"
                    SetConsoleColor colors.DefaultText
                    WriteLine "Erasing"
                    Do While Dir(App.path & "\.emr\backup", vbDirectory) <> ""
                        Call BackLine: Call ClearLine
                        WriteWord "Erasing " & IIf(GetTickCount Mod 4000 < 1000, ".", "") & IIf(GetTickCount Mod 4000 < 2000, ".", "") & IIf(GetTickCount Mod 4000 < 3000, ".", "")
                        Debuginfo.UpdateTimer_Timer
                        Sleep 10: DoEvents
                    Loop
                    SetConsoleColor colors.SuccessText
                    MkDir App.path & "\.emr\backup"
                    WriteWord "backup-Success: succeed in erasing all your backup ."
                    Set FSO3 = Nothing
                    Call ExitExec
                    Exit Sub
                Case "-list"
                    Dim backupF As String, backupI As Integer, FSO As Object
                    Set FSO = PoolCreateObject("Scripting.FileSystemObject")
                    backupF = Dir(App.path & "\.emr\backup\", vbDirectory)
                    Do While backupF <> ""
                        If backupF <> "." And backupF <> ".." Then
                            backupI = backupI + 1
                            WriteLine backupI & ". " & backupF & "  (" & FSO.GetFolder(App.path & "\.emr\backup\" & backupF).size & " bytes)"
                        End If
                        backupF = Dir(, vbDirectory)
                        'reserve : folder size
                    Loop
                    SetConsoleColor colors.ContentText
                    WriteWord backupI & " backups in total ."
                    Set FSO = Nothing
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
        Case "project"
            Select Case Params(1)
                Case "-help"
                    WriteLine "Set your current project ."
                    WriteLine "-reset : reset your version information ."
                    Call ExitExec
                    Exit Sub
                Case "-reset"
                    Open App.path & "\.emerald" For Output As #1
                    Print #1, 0 'version
                    Print #1, Now 'Update Time
                    Print #1, False
                    Close #1
                    SetConsoleColor colors.SuccessText
                    WriteWord "project-Success: succeed in reseting the version information ."
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
        Case "info"
            Select Case Params(1)
                Case "-help"
                    WriteLine "Show detail Emerald Information ."
                    Call ExitExec
                    Exit Sub
                Case ""
                    SetConsoleColor colors.DefaultText
                    WriteWord "Project name : "
                    SetConsoleColor colors.ContentText
                    WriteWord App.Title
                    NewLine
                    SetConsoleColor colors.DefaultText
                    WriteWord "Save state : "
                    SetConsoleColor IIf(ESave Is Nothing, colors.ErrText, colors.HighLightText)
                    WriteWord IIf(ESave Is Nothing, "[Nothing] ", "[Safe] ")
                    If Not ESave Is Nothing Then
                        SetConsoleColor colors.ContentText
                        WriteWord ESave.Count & " items in total ."
                        NewLine
                        SetConsoleColor colors.DefaultText
                        WriteWord "Save key : "
                        SetConsoleColor colors.HighLightText
                        WriteWord ESave.Key
                        NewLine
                        SetConsoleColor colors.DefaultText
                        WriteWord "Autosave : " & ESave.AutoSave
                    End If
                    NewLine
                    
                    If Abs(FPSct - 1000) > 60 Then
                        SetConsoleColor colors.ErrText
                        WriteLine "NOTE: It seems that you are using a timer to draw ."
                    End If
                    NewLine
                    
                    SetConsoleColor colors.DefaultText
                    WriteLine "Active page : " & ECore.ActivePage
                    WriteLine "FPS " & FPS & " , max FPS " & Int(1000 / Int(FPSct / FPS)) & " (per frame " & Int(FPSct / FPS) & " ms)"
                    WriteLine "Button " & Mouse.button & " , click state " & Mouse.State & " , in (" & Mouse.X & "," & Mouse.y & ")"
                    NewLine
                    SetConsoleColor colors.HighLightText
                    WriteLine "Visual Basic 6.0 IDE"
                    SetConsoleColor colors.DefaultText
                    WriteLine "ThreadID : " & VB.App.ThreadID
                    WriteLine "hInstance : " & VB.App.hInstance
                    WriteLine "CompanyName : " & VB.App.CompanyName
                    WriteLine "Form counts : " & VB.Forms.Count
                    SetConsoleColor colors.HighLightText
                    WriteLine "Screen Information"
                    SetConsoleColor colors.DefaultText
                    WriteLine "Size : " & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
                    WriteLine "DPI : " & 15 / Screen.TwipsPerPixelX * 96 & "(" & Int(15 / Screen.TwipsPerPixelX * 100) & "%)"
                    SetConsoleColor colors.HighLightText
                    WriteLine "Creation Information"
                    SetConsoleColor colors.DefaultText
                    WriteLine "Save : " & IIf(ESave Is Nothing, "[Missing]", ObjPtr(ESave))
                    WriteLine "Page Manager : " & IIf(ECore Is Nothing, "[Missing]", ObjPtr(ECore))
                    WriteLine "Font : " & IIf(EF Is Nothing, "[Missing]", ObjPtr(EF))
                    WriteLine "Animation : " & IIf(EAni Is Nothing, "[Missing]", ObjPtr(EAni))
                    WriteLine "Music List : " & IIf(EAni Is Nothing, "[Missing]", ObjPtr(EMusic))
                    WriteLine "Assets Tree : " & UBound(AssetsTrees) & " trees in total ."
                    For I = 1 To UBound(AssetsTrees)
                        WriteLine I & ". " & AssetsTrees(I).path & " (" & IIf(IsMissing(AssetsTrees(I).arg1), "[Missing]", AssetsTrees(I).arg1) & "," & IIf(IsMissing(AssetsTrees(I).arg2), "[Missing]", AssetsTrees(I).arg2) & ") , " & UBound(AssetsTrees(I).Files) & " files in total ."
                    Next
                    WriteLine "Current EditBox : " & TextHandle
                    WriteLine "Bass : " & BassInstalled
                    WriteLine "GDI+ : " & EmeraldInstalled
                    WriteLine LoadedCount & " images loaded (" & ReLoadCount & " reload ) ."
                    WriteLine "CrashMap : " & UBound(ColorLists)
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
        Case "page"
            Select Case Params(1)
                Case "-help"
                    WriteLine "Manage your pages ."
                    WriteLine "-change [name] : Change active page ."
                    WriteWord "-activepage : Show the name of active page ."
                    Call ExitExec
                    Exit Sub
                Case "-change"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    ECore.ActivePage = Params(2)
                    SetConsoleColor colors.SuccessText
                    WriteWord "page-Success: change success ."
                    Call ExitExec
                    Exit Sub
                Case "-activepage"
                    WriteWord "Active page : " & ECore.ActivePage
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
        Case "save"
            If ESave Is Nothing Then
                SetConsoleColor colors.ErrText
                WriteWord "save-Error: create your save first ."
                Call ExitExec
                Exit Sub
            End If
            Select Case Params(1)
                Case "-help"
                    WriteLine "Edit your save ."
                    WriteLine "-remove [name] : Clear your data ."
                    WriteLine "-add [name] [value]: Add new data ."
                    WriteLine "-edit [name] [value]: Edit your data ."
                    WriteLine "-key : Show your save key ."
                    WriteLine "-changekey [key] : Change your save key <danger> ."
                    WriteLine "-getkey : Get a random key ."
                    WriteLine "-save : Write your save out to the file right now ."
                    WriteLine "-saveswitch [boolean]: Switch the setting 'AutoSave' ."
                    WriteLine "-count : Get the count of your data ."
                    WriteWord "-list : List your data ."
                    Call ExitExec
                    Exit Sub
                Case "-changekey"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    SetConsoleColor colors.ErrText
                    WriteLine "Danger !"
                    WriteLine "The consequences are at one's own expense !!!"
                    WriteLine "Your save will may be lost !"
                    SetConsoleColor colors.DefaultText
                    WriteLine "New key : " & Params(2)
                    WriteLine "Change your save key any way ? [Y/N]"
                    AllowInput
                    Params(1) = "-changekey2"
                Case "-changekey2"
                Case "-changekey3"
                    ESave.Key = Params(2)
                    SetConsoleColor colors.ContentText
                    For I = 1 To ESave.Count
                        ESave.PutData ESave.GetName(Val(I)), ESave.GetData(ESave.GetName(Val(I)))
                        WriteLine "Rewriting : " & ESave.GetName(Val(I)) & "(item " & I & ")"
                    Next
                    SetConsoleColor colors.ErrText
                    WriteWord "NOTE : Remember to change your key in your code !"
                    Call ExitExec
                    Exit Sub
                Case "-save"
                    ESave.Save
                    SetConsoleColor colors.SuccessText
                    WriteWord "save-Success:saved ."
                    Call ExitExec
                    Exit Sub
                Case "-count"
                    SetConsoleColor colors.DefaultText
                    WriteWord ESave.Count & " items in total ."
                    Call ExitExec
                    Exit Sub
                Case "-getkey"
                    SetConsoleColor colors.DefaultText
                    WriteWord GetBMKey
                    Call ExitExec
                    Exit Sub
                Case "-list"
                    SetConsoleColor colors.DefaultText
                    WriteLine ESave.Count & " items in total ."
                    Dim savestr As String
                    For I = 1 To ESave.Count
                        savestr = ESave.GetData(ESave.GetName(Val(I)))
                        SetConsoleColor IIf(savestr = "", colors.ContentText, colors.DefaultText)
                        WriteLine I & ". " & IIf(savestr = "", "[Removed]", "[Safe]") & "    " & ESave.GetName(Val(I)) & "    " & savestr
                    Next
                    Call ExitExec
                    Exit Sub
                Case "-key"
                    SetConsoleColor colors.DefaultText
                    WriteLine "Save key : " & ESave.Key
                    SetConsoleColor colors.ErrText
                    WriteWord "NOTE : Please keep your key properly ."
                    Call ExitExec
                    Exit Sub
                Case "-remove"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    ESave.PutData Params(2), ""
                    SetConsoleColor colors.SuccessText
                    WriteWord "save-Success: succeed in removing data ."
                    Call ExitExec
                Case "-add"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    If UBound(Params) < 3 Then MissingParam = 3: GoTo MissingParams
                    ESave.PutData Params(2), Params(3)
                    SetConsoleColor colors.SuccessText
                    WriteWord "save-Success: succeed in adding data ."
                    Call ExitExec
                Case "-edit"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    If UBound(Params) < 3 Then MissingParam = 3: GoTo MissingParams
                    ESave.PutData Params(2), Params(3)
                    SetConsoleColor colors.SuccessText
                    WriteWord "save-Success: succeed in writting data ."
                    Call ExitExec
                Case "-saveswitch"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "true" Then
                        ESave.AutoSave = True
                        SetConsoleColor colors.SuccessText
                        WriteWord "save-Success: AutoSave " & ESave.AutoSave
                        Call ExitExec
                    ElseIf Params(2) = "false" Then
                        ESave.AutoSave = False
                        SetConsoleColor colors.SuccessText
                        WriteWord "save-Success: AutoSave " & ESave.AutoSave
                        Call ExitExec
                    Else
                        SetConsoleColor colors.ErrText
                        WriteWord "save-Error: use 'true' or 'false' instead ."
                        Call ExitExec
                    End If
                Case Else: GoTo UnKnownCmd
            End Select
        Case "checkupdate"
            Select Case Params(1)
                Case "-help"
                    WriteLine "Check online update for Emerald ."
                    WriteWord "-force : Ignore update settings ."
                    Call ExitExec
                    Exit Sub
                Case ""
                    Params(1) = "-wait"
                    Call CheckUpdate(False)
                    Call ExitExec
                Case "-force"
                    Params(1) = "-wait"
                    Call CheckUpdate(True)
                    Call ExitExec
                Case "-wait"
                Case Else: GoTo UnKnownCmd
            End Select
        Case "switchdebug"
            Select Case Params(1)
                Case "-help"
                    WriteWord "[index] [boolean] : Switch the debug function ."
                    Call ExitExec
                    Exit Sub
                Case Else
                    If UBound(Params) < 1 Then MissingParam = 1: GoTo MissingParams
                    If Val(Params(1)) >= 1 And Val(Params(1)) <= 3 Then
                        If UBound(Params) < 2 Then
                            MissingParam = 2: GoTo MissingParams
                        End If
                        If Params(2) = "true" Then
                            Select Case Val(Params(1))
                                Case 1: Debug_data = False
                                Case 2: Debug_pos = False
                                Case 3: Debug_focus = False
                            End Select
                            DebugWindow.touchArea_Click (3 - Val(Params(1))) + 3
                            SetConsoleColor colors.SuccessText
                            WriteWord "switchdebug-Success: alright ."
                            Call ExitExec
                        ElseIf Params(2) = "false" Then
                            Select Case Val(Params(1))
                                Case 1: Debug_data = True
                                Case 2: Debug_pos = True
                                Case 3: Debug_focus = True
                            End Select
                            DebugWindow.touchArea_Click (3 - Val(Params(1))) + 3
                            SetConsoleColor colors.SuccessText
                            WriteWord "switchdebug-Success: alright ."
                            Call ExitExec
                        Else
                            SetConsoleColor colors.ErrText
                            WriteWord "switchdebug-Error: use 'true' or 'false' instead ."
                            Call ExitExec
                        End If
                    ElseIf Params(1) = CStr(Val(Params(1))) Then
                        SetConsoleColor colors.ErrText
                        WriteWord "switchdebug-Error: the index must be between 1 and 3 ."
                        Call ExitExec
                    Else
                        GoTo UnKnownCmd
                    End If
            End Select
        Case "showfps"
            Call BackLine: Call ClearLine
            SetConsoleColor colors.DefaultText
            Select Case Params(1)
                Case "-detail":  WriteWord "FPS " & FPS & " , max FPS " & Int(1000 / Int(FPSct / FPS)) & " (per frame " & Int(FPSct / FPS) & " ms)"
                Case "-nomax": WriteWord "FPS " & FPS & " ."
                Case "-cost": WriteWord "Total " & FPSct & " ms , per frame " & Int(FPSct / FPS) & " ms ."
                Case "": WriteWord "FPS " & FPS & " , max FPS " & Int(1000 / Int(FPSct / FPS)) & "."
                Case "-help"
                    WriteLine "Show informations about FPS ."
                    WriteLine "-nomax : Show FPS without max FPS ."
                    WriteLine "-cost : Show how much time the drawing cost ."
                    WriteWord "-detail : Show the detail information ."
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
            Call NextLine: Call ClearLine
            WriteWord "Press [x] to exit ."
        Case "mousestate"
            Call BackLine: Call ClearLine
            SetConsoleColor colors.DefaultText
            Select Case Params(1)
                Case "": WriteWord "Button " & Mouse.button & " , click state " & Mouse.State & " , in (" & Mouse.X & "," & Mouse.y & ")"
                Case "-help"
                    WriteWord "No any commands ."
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
            Call NextLine: Call ClearLine
            WriteWord "Press [x] to exit ."
        Case Else
            SetConsoleColor colors.DefaultText
            WriteWord "Error: No such a command as '"
            SetConsoleColor colors.HighLightText
            WriteWord Exec
            SetConsoleColor colors.DefaultText
            WriteWord "' ."
            Call ExitExec
            Exit Sub
    End Select
    
    Exit Sub
    
UnKnownCmd:
    SetConsoleColor colors.ErrText
    WriteWord "Error: Unknown command '" & Params(1) & "' , type '" & Exec & " -help' for help ."
    Call ExitExec
    Exit Sub
    
MissingParams:
    SetConsoleColor colors.ErrText
    WriteWord Exec & "-Error: param " & MissingParam & " not found , type '" & Exec & " -help' for help ."
    Call ExitExec
End Sub
Public Sub AllowInput()
    InputAllow = True
    InputingText = ""
End Sub
Public Sub AllowCmd()
    Call AllowInput
    SetConsoleColor colors.SuccessText
    WriteWord "Emr>"
    SetConsoleColor colors.ErrText
    WriteWord "VB>"
    SetConsoleColor colors.HighLightText
    WriteWord App.Title & ">"
    If Exec <> "" Then
        SetConsoleColor colors.DefaultText
        WriteWord Exec & ">"
    End If
End Sub
Public Sub BackLine()
    CuY = CuY - 25
    CuX = 0
End Sub
Public Sub NextLine()
    CuY = CuY + 25
    CuX = 0
End Sub
Public Sub ClearLine()
    Dim ODC As Long, OGG As Long
    ODC = Page.CDC: OGG = Page.GG
    
    Page.CDC = ConsoleDC: Page.GG = ConsoleGG
    
    GdipSetClipRect ConsoleGG, 0, CuY - 1, GW - 40, 22, CombineModeReplace
    GdipGraphicsClear ConsoleGG, argb(0, 0, 0, 0)
    GdipResetClip ConsoleGG
    
    Page.CDC = ODC: Page.GG = OGG
End Sub
Public Sub WriteWord(txt As String)
    Dim ODC As Long, OGG As Long
    ODC = Page.CDC: OGG = Page.GG
    
    Page.CDC = ConsoleDC: Page.GG = ConsoleGG
    
    Page.Writes txt, CuX, CuY, 16, Color
    
    CuX = CuX + EF.GetWidth(ConsoleGG, txt, 16, StringAlignmentNear, FontStyleRegular)
    
    Page.CDC = ODC: Page.GG = OGG
End Sub
Public Sub WriteLine(txt As String)
    Call WriteWord(txt)
    Call NewLine
End Sub
Public Sub NewLine()
    CuY = CuY + 25
    If CuY + 25 > 3222 Then
        CuY = CuY - 25
        BitBlt ConsoleDC, 0, 0, GW - 20, 3222 - 25, ConsoleDC, 0, 25, vbSrcCopy
        ClearLine
    End If
    If CuY - sy * 3222 >= GH - 80 Then
        sy = (CuY - GH + 80) / 3222
    End If
    CuX = 0
End Sub
Public Sub SetConsoleColor(C As Long)
    Color = C
End Sub
Public Sub InitConsole()
    ConsoleDC = CreateCDC(GW - 20, 3222)
    PoolCreateFromHdc ConsoleDC, ConsoleGG
    GdipSetSmoothingMode ConsoleGG, SmoothingModeAntiAlias
    GdipSetTextRenderingHint ConsoleGG, TextRenderingHintAntiAlias
End Sub
Public Sub Update2()
    Page.Clear argb(220, 16, 16, 16)
    
    If FirstDisplay = False Then
        NewLine
        SetConsoleColor colors.DefaultText
        WriteWord "Emerald Console Version "
        SetConsoleColor colors.HighLightText
        WriteLine Version
        SetConsoleColor colors.ContentText
        WriteLine "Copyright (C) 2019 Error404 all rights reserved ."
        NewLine
        SetConsoleColor colors.DefaultText
        WriteWord "Project name : "
        SetConsoleColor colors.ContentText
        WriteWord App.Title
        NewLine
        SetConsoleColor colors.DefaultText
        WriteWord "Save state : "
        SetConsoleColor IIf(ESave Is Nothing, colors.ErrText, colors.HighLightText)
        WriteWord IIf(ESave Is Nothing, "[Nothing] ", "[Safe] ")
        If Not ESave Is Nothing Then
            SetConsoleColor colors.ContentText
            WriteWord ESave.Count & " items in total ."
            NewLine
            SetConsoleColor colors.DefaultText
            WriteWord "Save key : "
            SetConsoleColor colors.HighLightText
            WriteWord ESave.Key
        End If
        NewLine
        
        If Abs(FPSct - 1000) > 60 Then
            SetConsoleColor colors.ErrText
            WriteLine "NOTE: It seems that you are using a timer to draw ."
        End If
        NewLine
        AllowCmd
        FirstDisplay = True
    End If
    
    Call RunExec
    
    PaintDC ConsoleDC, Page.CDC, 20, 40, 0, sy * 3222, GW - 40, IIf(3222 - sy * 3222 < GH, 3222 - sy * 3222, GH), 1
    
    Dim pro As Long, alpha As Single

    pro = GetTickCount Mod 1000
    If pro <= 700 Then
        alpha = 1 - Cubic(pro / 700, 0, 1, 1, 1)
    Else
        alpha = Cubic((pro - 700) / 300, 0, 1, 1, 1)
    End If
    Page.Writes InputingText, CuX + 20, CuY - sy * 3222 + 40, 16, argb(255, 242, 242, 242)
    
    If GetActiveWindow = Debuginfo.Hwnd Then Page.Paint 0, CuX + 20 + EF.GetWidth(Page.GG, InputingText, 16, StringAlignmentNear, FontStyleRegular), CuY - sy * 3222 + 40, 3, 20, argb(alpha * 255, 255, 255, 255)
    
    If CuY + 80 >= GH Then
        Dim MaxY As Single
        MaxY = (CuY - GH + 80) / 3220
        Page.Paint 2, GW - 17, 60 + (sy / MaxY) * (GH - 60 - 20 - (80 - CuY / 3222 * 60)), 4, 80 - CuY / 3222 * 60, argb(255, 128, 128, 128), 20
        NeedScroll = True
    End If
    
    Page.Paint 0, 1, 1, GW - 1, GH - 1, argb(255, 0, 183, 195), size:=2, style:=1
    Page.Paint 0, 0, 0, GW + 1, 40, argb(255, 0, 183, 195)
    Page.Writes "Emerald Console", 20, 9, 16, argb(255, 27, 27, 27)
    
End Sub
Public Sub Update()
    If PageMark = 1 Then Call Update2: Exit Sub

    On Error Resume Next
    Dim W As Long, H As Long, pro As Single
    W = GW / 21
    Dim LY(20) As Long
    For I = 0 To 20
        H = (FPSRecord(I) - 24) / 40 * (GH - 7)
        If H <= 0 Then H = 1
        If H > GH - 7 Then H = GH - 7
        LY(I) = GH - H
        If I > 0 Then
            pro = Cubic(H / GH, 0, 1, 1, 1)
            Page.PaintLine I * W - W, LY(I - 1), I * W, LY(I), argb(pro * 40 + 20, 255, 255, 255), 2
        End If
    Next
    
    If FPS <> 0 Then
        Dim TColor As Long
        Select Case FPS
            Case Is >= 60: TColor = argb(255, 9, 241, 117)
            Case Is >= 32: TColor = argb(255, 0, 176, 240)
            Case Is >= 24: TColor = argb(255, 253, 86, 60)
            Case Else: TColor = argb(255, 255, 0, 0)
        End Select
        Page.Writes "FPS:" & FPS & "/" & Int(1000 / Int(FPSct / FPS)), 55, 16, 18, TColor, 120
    End If
    
    Page.DrawImage "profile.png", 5, 54 / 2 - 48 / 2 + 1, alpha:=1
    
    If Debug_mouse Then Page.Paint 0, GW - 48, 0, 49, GH, argb(60, 255, 255, 255)
    Page.DrawImage "mouse" & Mouse.button & ".png", GW - 48, 54 / 2 - 48 / 2 + 1, alpha:=1

    If Debuginfo.Visible Then Page.Paint 0, GW - 48 * 3, 0, 49, GH, argb(60, 255, 255, 255)
    Page.DrawImage "menu.png", GW - 48 * 3, 54 / 2 - 48 / 2 + 1, alpha:=1

    If Debug_umode Then Page.Paint 0, GW - 48 * 5, 0, 49, GH, argb(60, 255, 255, 255)
    Page.DrawImage "u" & Debug_umode & ".png", GW - 48 * 5, 54 / 2 - 48 / 2 + 1, alpha:=1

    If Debug_pos Then Page.Paint 0, GW - 48 * 4, 0, 49, GH, argb(60, 255, 255, 255)
    Page.DrawImage "pos.png", GW - 48 * 4, 54 / 2 - 48 / 2 + 1, alpha:=1

    If Debug_data Then Page.Paint 0, GW - 48 * 2, 0, 49, GH, argb(60, 255, 255, 255)
    Page.DrawImage "data.png", GW - 48 * 2, 54 / 2 - 48 / 2 + 1, alpha:=1

    Page.Paint 0, 0, 0, GW + 1, 6, argb(255, 36, 173, 243)
    If FPS <> 0 Then
        Page.Paint 0, 0, 0, FPS / Int(1000 / Int(FPSct / FPS)) * (GW + 1), 6, argb(255, 130, 221, 255)
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)

End Sub
Public Sub Enter()

End Sub
Public Sub Leave()

End Sub
Public Sub Gotfocus()

End Sub
Public Sub Lostfocus()

End Sub
Public Sub AnimationDone(id As Integer)

End Sub


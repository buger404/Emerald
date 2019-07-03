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

Public Page As GPage, GW As Long, GH As Long
Public PageMark As Long

Public InputingText As String, InputAllow As Boolean

Private Enum Colors
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

Public SY As Single, NeedScroll As Boolean
Public CuY As Long, CuX As Long
Public Sub EndInput()
    ChooseCmd = CurrentCmd
    SetConsoleColor Colors.DefaultText
    WriteWord InputingText
    NewLine
End Sub
Public Sub ApplyCmd()
    Call EndInput
    Select Case Exec
        Case ""
            Dim Param() As String
            Param = Split(InputingText, " ")
            If UBound(Param) = 0 Then ReDim Preserve Param(1)
            
            If InputingText = "" Or Exec <> "" Then GoTo SkipRecord
            
            If CurrentCmd = UBound(HistoryCmd) Then
                For i = 1 To UBound(HistoryCmd)
                    HistoryCmd(i - 1) = HistoryCmd(i)
                Next
            End If
            
            HistoryCmd(CurrentCmd) = InputingText
            If CurrentCmd < UBound(HistoryCmd) Then CurrentCmd = CurrentCmd + 1
    
SkipRecord:

            InputAllow = False
            If InputingText <> "" Then
                Exec = Param(0)
                Params = Param
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
        Case ""
            If Key = vbKeyUp Then
                ChooseCmd = ChooseCmd - 1
                If ChooseCmd < 0 Then ChooseCmd = 0
                InputingText = HistoryCmd(ChooseCmd)
            End If
            If Key = vbKeyDown Then
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
        SetConsoleColor Colors.ErrText
        WriteLine "Emerald-Error: connect to the Internet first ."
        Exit Sub
    End If
    
    Dim data As New GSaving
    data.Create "Emerald.Core"
    data.AutoSave = True
    If Now - CDate(data.GetData("UpdateTime")) >= UpdateCheckInterval Or force Or data.GetData("UpdateAble") = 1 Then
        data.PutData "UpdateTime", Now
        
        Dim xmlHttp As Object, Ret As String, Start As Long
        
        SetConsoleColor Colors.DefaultText
        WriteLine "Emerald: connecting to the server"
        Start = GetTickCount
        Do While GetTickCount - Start <= 2000
            Call BackLine: Call ClearLine
            WriteLine "Emerald: connecting to the server" & IIf(GetTickCount Mod 1000 > 250, ".", "") & IIf(GetTickCount Mod 1000 > 500, ".", "") & IIf(GetTickCount Mod 1000 > 750, ".", "")
            Debuginfo.UpdateTimer_Timer
            Sleep 10: DoEvents
        Loop
        
        Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
        xmlHttp.Open "GET", "https://raw.githubusercontent.com/Red-Error404/Emerald/master/Version.txt", True
        xmlHttp.send
                     
        Start = GetTickCount
        
        Do While xmlHttp.ReadyState <> 4
            If GetTickCount - Start >= UpdateTimeOut Then
                SetConsoleColor Colors.ErrText
                WriteLine "Emerald-Error: failed to connect , time out ."
                Exit Sub
            End If
            Call BackLine: Call ClearLine
            WriteLine "Emerald: connecting to the server" & IIf(GetTickCount Mod 4000 < 1000, ".", "") & IIf(GetTickCount Mod 4000 < 2000, ".", "") & IIf(GetTickCount Mod 4000 < 3000, ".", "")
            Debuginfo.UpdateTimer_Timer
            Sleep 10: DoEvents
        Loop
        Ret = xmlHttp.responseText
        Set xmlHttp = Nothing
        SetConsoleColor Colors.SuccessText
        WriteLine "Emerald-Success: the lastest version from the server " & Val(Ret)
        SetConsoleColor Colors.ContentText
        WriteLine "Emerald: current version " & Version
        
        If Val(Ret) > Version And App.LogMode = 0 Then
            SetConsoleColor Colors.SuccessText
            WriteLine "Emerald: update available , plz visit 'https://github.com/Red-Error404/Emerald/release' to download ."
            Exit Sub
        End If
    Else
        SetConsoleColor Colors.HighLightText
        WriteLine "Emerald-Skip: update has been checked at " & CDate(data.GetData("UpdateTime")) & "."
    End If
    
    Set data = Nothing
End Sub
Public Sub RunExec()
    If Exec = "" Then Exit Sub

    Dim MissingParam As Integer

    Select Case Exec
        Case "info"
            Select Case Params(1)
                Case "-help"
                    WriteLine "Show detail Emerald Information ."
                    Call ExitExec
                    Exit Sub
                Case ""
                    SetConsoleColor Colors.DefaultText
                    WriteWord "Project name : "
                    SetConsoleColor Colors.ContentText
                    WriteWord App.Title
                    NewLine
                    SetConsoleColor Colors.DefaultText
                    WriteWord "Save state : "
                    SetConsoleColor IIf(ESave Is Nothing, Colors.ErrText, Colors.HighLightText)
                    WriteWord IIf(ESave Is Nothing, "[Nothing] ", "[Safe] ")
                    If Not ESave Is Nothing Then
                        SetConsoleColor Colors.ContentText
                        WriteWord ESave.Count & " items in total ."
                        NewLine
                        SetConsoleColor Colors.DefaultText
                        WriteWord "Save key : "
                        SetConsoleColor Colors.HighLightText
                        WriteWord ESave.Key
                        NewLine
                        SetConsoleColor Colors.DefaultText
                        WriteWord "Autosave : " & ESave.AutoSave
                    End If
                    NewLine
                    
                    If Abs(FPSct - 1000) > 60 Then
                        SetConsoleColor Colors.ErrText
                        WriteLine "NOTE: It seems that you are using a timer to draw ."
                    End If
                    NewLine
                    
                    SetConsoleColor Colors.DefaultText
                    WriteLine "Active page : " & ECore.ActivePage
                    WriteLine "FPS " & FPS & " , max FPS " & Int(1000 / Int(FPSct / FPS)) & " (per frame " & Int(FPSct / FPS) & " ms)"
                    WriteLine "Button " & Mouse.button & " , click state " & Mouse.state & " , in (" & Mouse.x & "," & Mouse.y & ")"
                    NewLine
                    SetConsoleColor Colors.HighLightText
                    WriteLine "Visual Basic 6.0 IDE"
                    SetConsoleColor Colors.DefaultText
                    WriteLine "ThreadID : " & VB.App.ThreadID
                    WriteLine "hInstance : " & VB.App.hInstance
                    WriteLine "CompanyName : " & VB.App.CompanyName
                    WriteLine "Form counts : " & VB.Forms.Count
                    SetConsoleColor Colors.HighLightText
                    WriteLine "Screen Information"
                    SetConsoleColor Colors.DefaultText
                    WriteLine "Size : " & Screen.Width / Screen.TwipsPerPixelX & "x" & Screen.Height / Screen.TwipsPerPixelY
                    WriteLine "DPI : " & 15 / Screen.TwipsPerPixelX * 96 & "(" & Int(15 / Screen.TwipsPerPixelX * 100) & "%)"
                    SetConsoleColor Colors.HighLightText
                    WriteLine "Creation Information"
                    SetConsoleColor Colors.DefaultText
                    WriteLine "Save : " & IIf(ESave Is Nothing, "[Missing]", ObjPtr(ESave))
                    WriteLine "Page Manager : " & IIf(ECore Is Nothing, "[Missing]", ObjPtr(ECore))
                    WriteLine "Font : " & IIf(EF Is Nothing, "[Missing]", ObjPtr(EF))
                    WriteLine "Animation : " & IIf(EAni Is Nothing, "[Missing]", ObjPtr(EAni))
                    WriteLine "Music List : " & IIf(EAni Is Nothing, "[Missing]", ObjPtr(EMusic))
                    WriteLine "Assets Tree : " & UBound(AssetsTrees) & " trees in total ."
                    For i = 1 To UBound(AssetsTrees)
                        WriteLine i & ". " & AssetsTrees(i).path & " (" & AssetsTrees(i).arg1 & "," & AssetsTrees(i).arg2 & ") , " & UBound(AssetsTrees(i).files) & " files in total ."
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
                    SetConsoleColor Colors.SuccessText
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
                SetConsoleColor Colors.ErrText
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
                    SetConsoleColor Colors.ErrText
                    WriteLine "Danger !"
                    WriteLine "The consequences are at one's own expense !!!"
                    WriteLine "Your save will may be lost !"
                    SetConsoleColor Colors.DefaultText
                    WriteLine "New key : " & Params(2)
                    WriteLine "Change your save key any way ? [Y/N]"
                    AllowInput
                    Params(1) = "-changekey2"
                Case "-changekey2"
                Case "-changekey3"
                    ESave.Key = Params(2)
                    SetConsoleColor Colors.ContentText
                    For i = 1 To ESave.Count
                        ESave.PutData ESave.GetName(Val(i)), ESave.GetData(ESave.GetName(Val(i)))
                        WriteLine "Rewriting : " & ESave.GetName(Val(i)) & "(item " & i & ")"
                    Next
                    SetConsoleColor Colors.ErrText
                    WriteWord "NOTE : Remember to change your key in your code !"
                    Call ExitExec
                    Exit Sub
                Case "-save"
                    ESave.Save
                    SetConsoleColor Colors.SuccessText
                    WriteWord "save-Success:saved ."
                    Call ExitExec
                    Exit Sub
                Case "-count"
                    SetConsoleColor Colors.DefaultText
                    WriteWord ESave.Count & " items in total ."
                    Call ExitExec
                    Exit Sub
                Case "-getkey"
                    SetConsoleColor Colors.DefaultText
                    WriteWord GetBMKey
                    Call ExitExec
                    Exit Sub
                Case "-list"
                    SetConsoleColor Colors.DefaultText
                    WriteLine ESave.Count & " items in total ."
                    Dim savestr As String
                    For i = 1 To ESave.Count
                        savestr = ESave.GetData(ESave.GetName(Val(i)))
                        SetConsoleColor IIf(savestr = "", Colors.ContentText, Colors.DefaultText)
                        WriteLine i & ". " & IIf(savestr = "", "[Removed]", "[Safe]") & "    " & ESave.GetName(Val(i)) & "    " & savestr
                    Next
                    Call ExitExec
                    Exit Sub
                Case "-key"
                    SetConsoleColor Colors.DefaultText
                    WriteLine "Save key : " & ESave.Key
                    SetConsoleColor Colors.ErrText
                    WriteWord "NOTE : Please keep your key properly ."
                    Call ExitExec
                    Exit Sub
                Case "-remove"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    ESave.PutData Params(2), ""
                    SetConsoleColor Colors.SuccessText
                    WriteWord "save-Success: succeed in removing data ."
                    Call ExitExec
                Case "-add"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    If UBound(Params) < 3 Then MissingParam = 3: GoTo MissingParams
                    ESave.PutData Params(2), Params(3)
                    SetConsoleColor Colors.SuccessText
                    WriteWord "save-Success: succeed in adding data ."
                    Call ExitExec
                Case "-edit"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    If UBound(Params) < 3 Then MissingParam = 3: GoTo MissingParams
                    ESave.PutData Params(2), Params(3)
                    SetConsoleColor Colors.SuccessText
                    WriteWord "save-Success: succeed in writting data ."
                    Call ExitExec
                Case "-saveswitch"
                    If UBound(Params) < 2 Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "" Then MissingParam = 2: GoTo MissingParams
                    If Params(2) = "true" Then
                        ESave.AutoSave = True
                        SetConsoleColor Colors.SuccessText
                        WriteWord "save-Success: AutoSave " & ESave.AutoSave
                        Call ExitExec
                    ElseIf Params(2) = "false" Then
                        ESave.AutoSave = False
                        SetConsoleColor Colors.SuccessText
                        WriteWord "save-Success: AutoSave " & ESave.AutoSave
                        Call ExitExec
                    Else
                        SetConsoleColor Colors.ErrText
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
                            SetConsoleColor Colors.SuccessText
                            WriteWord "switchdebug-Success: alright ."
                            Call ExitExec
                        ElseIf Params(2) = "false" Then
                            Select Case Val(Params(1))
                                Case 1: Debug_data = True
                                Case 2: Debug_pos = True
                                Case 3: Debug_focus = True
                            End Select
                            DebugWindow.touchArea_Click (3 - Val(Params(1))) + 3
                            SetConsoleColor Colors.SuccessText
                            WriteWord "switchdebug-Success: alright ."
                            Call ExitExec
                        Else
                            SetConsoleColor Colors.ErrText
                            WriteWord "switchdebug-Error: use 'true' or 'false' instead ."
                            Call ExitExec
                        End If
                    ElseIf Params(1) = CStr(Val(Params(1))) Then
                        SetConsoleColor Colors.ErrText
                        WriteWord "switchdebug-Error: the index must be between 1 and 3 ."
                        Call ExitExec
                    Else
                        GoTo UnKnownCmd
                    End If
            End Select
        Case "showfps"
            Call BackLine: Call ClearLine
            SetConsoleColor Colors.DefaultText
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
            SetConsoleColor Colors.DefaultText
            Select Case Params(1)
                Case "": WriteWord "Button " & Mouse.button & " , click state " & Mouse.state & " , in (" & Mouse.x & "," & Mouse.y & ")"
                Case "-help"
                    WriteWord "No any commands ."
                    Call ExitExec
                    Exit Sub
                Case Else: GoTo UnKnownCmd
            End Select
            Call NextLine: Call ClearLine
            WriteWord "Press [x] to exit ."
        Case Else
            SetConsoleColor Colors.DefaultText
            WriteWord "Error: No such a command as '"
            SetConsoleColor Colors.HighLightText
            WriteWord Exec
            SetConsoleColor Colors.DefaultText
            WriteWord "' ."
            NewLine
            Call AllowCmd
            Exec = ""
            Exit Sub
    End Select
    
    Exit Sub
    
UnKnownCmd:
    SetConsoleColor Colors.ErrText
    WriteWord "Error: Unknown command '" & Params(1) & "' , type '" & Exec & " -help' for help ."
    Call ExitExec
    Exit Sub
    
MissingParams:
    SetConsoleColor Colors.ErrText
    WriteWord Exec & "-Error: param " & MissingParam & " not found , type '" & Exec & " -help' for help ."
    Call ExitExec
End Sub
Public Sub AllowInput()
    InputAllow = True
    InputingText = ""
End Sub
Public Sub AllowCmd()
    Call AllowInput
    SetConsoleColor Colors.SuccessText
    WriteWord "Emr>"
    SetConsoleColor Colors.ErrText
    WriteWord "VB>"
    SetConsoleColor Colors.HighLightText
    WriteWord App.Title & ">"
    If Exec <> "" Then
        SetConsoleColor Colors.DefaultText
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
    If CuY - SY * 3222 >= GH - 80 Then
        SY = (CuY - GH + 80) / 3222
    End If
    CuX = 0
End Sub
Public Sub SetConsoleColor(c As Long)
    Color = c
End Sub
Public Sub InitConsole()
    ConsoleDC = CreateCDC(GW - 20, 3222)
    GdipCreateFromHDC ConsoleDC, ConsoleGG
    GdipSetSmoothingMode ConsoleGG, SmoothingModeAntiAlias
    GdipSetTextRenderingHint ConsoleGG, TextRenderingHintAntiAlias
End Sub
Public Sub Update2()
    Page.Clear argb(220, 16, 16, 16)
    
    If FirstDisplay = False Then
        NewLine
        SetConsoleColor Colors.DefaultText
        WriteWord "Emerald Console Version "
        SetConsoleColor Colors.HighLightText
        WriteLine Version
        SetConsoleColor Colors.ContentText
        WriteLine "Copyright (C) 2019 Error404 all rights reserved ."
        NewLine
        SetConsoleColor Colors.DefaultText
        WriteWord "Project name : "
        SetConsoleColor Colors.ContentText
        WriteWord App.Title
        NewLine
        SetConsoleColor Colors.DefaultText
        WriteWord "Save state : "
        SetConsoleColor IIf(ESave Is Nothing, Colors.ErrText, Colors.HighLightText)
        WriteWord IIf(ESave Is Nothing, "[Nothing] ", "[Safe] ")
        If Not ESave Is Nothing Then
            SetConsoleColor Colors.ContentText
            WriteWord ESave.Count & " items in total ."
            NewLine
            SetConsoleColor Colors.DefaultText
            WriteWord "Save key : "
            SetConsoleColor Colors.HighLightText
            WriteWord ESave.Key
        End If
        NewLine
        
        If Abs(FPSct - 1000) > 60 Then
            SetConsoleColor Colors.ErrText
            WriteLine "NOTE: It seems that you are using a timer to draw ."
        End If
        NewLine
        AllowCmd
        FirstDisplay = True
    End If
    
    Call RunExec
    
    PaintDC ConsoleDC, Page.CDC, 20, 40, 0, SY * 3222, GW - 40, IIf(3222 - SY * 3222 < GH, 3222 - SY * 3222, GH), 1
    
    Dim pro As Long, alpha As Single

    pro = GetTickCount Mod 1000
    If pro <= 700 Then
        alpha = 1 - Cubic(pro / 700, 0, 1, 1, 1)
    Else
        alpha = Cubic((pro - 700) / 300, 0, 1, 1, 1)
    End If
    Page.Writes InputingText, CuX + 20, CuY - SY * 3222 + 40, 16, argb(255, 242, 242, 242)
    
    If GetActiveWindow = Debuginfo.Hwnd Then Page.Paint 0, CuX + 20 + EF.GetWidth(Page.GG, InputingText, 16, StringAlignmentNear, FontStyleRegular), CuY - SY * 3222 + 40, 3, 20, argb(alpha * 255, 255, 255, 255)
    
    If CuY + 80 >= GH Then
        Dim MaxY As Single
        MaxY = (CuY - GH + 80) / 3220
        Page.Paint 2, GW - 17, 60 + (SY / MaxY) * (GH - 60 - 20 - (80 - CuY / 3222 * 60)), 4, 80 - CuY / 3222 * 60, argb(255, 128, 128, 128), 20
        NeedScroll = True
    End If
    
    Page.Paint 0, 0, 0, GW + 1, 40, argb(255, 0, 183, 195)
    Page.Writes "Emerald Console", 20, 9, 16, argb(255, 27, 27, 27)
    
End Sub
Public Sub Update()
    If PageMark = 1 Then Call Update2: Exit Sub

    On Error Resume Next
    
    Page.Writes FPS & "/" & Int(1000 / Int(FPSct / FPS)), 15, 14, 34, argb(205, 182, 174, 173)
    
    Page.DrawImage "mouse" & Mouse.button & ".png", GW - 138, 78 / 2 - 64 / 2, alpha:=1
    Page.DrawImage "menu.png", GW - 74, 78 / 2 - 64 / 2, alpha:=1
    If Debuginfo.Visible Then Page.Paint 0, GW - 74, 0, 64, 79, argb(50, 128, 128, 128)
    
    Page.DrawImage "focus.png", GW - 202, 78 / 2 - 64 / 2, alpha:=1
    If Debug_focus Then Page.Paint 0, GW - 202, 0, 64, 79, argb(50, 128, 128, 128)
    
    Page.DrawImage "pos.png", GW - 202 - 64, 78 / 2 - 64 / 2, alpha:=1
    If Debug_pos Then Page.Paint 0, GW - 202 - 64, 0, 64, 79, argb(50, 128, 128, 128)
    
    Page.DrawImage "data.png", GW - 202 - 128, 78 / 2 - 64 / 2, alpha:=1
    If Debug_data Then Page.Paint 0, GW - 202 - 128, 0, 64, 79, argb(50, 128, 128, 128)
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


Attribute VB_Name = "Process"
Sub Main()
    If Command$ <> "" Then
        Dim appn As String, f As String, t As String, p As String
        p = Replace(Command$, """", "")

        appn = InputBox("Input your app name.", "New Emerald Project")
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
        
        If Dir(p & "\core", vbDirectory) = "" Then MkDir p & "\core"
        CopyInto App.Path & "\core", p & "\core"
        If Dir(p & "\assets", vbDirectory) = "" Then MkDir p & "\assets"
        If Dir(p & "\assets\debug", vbDirectory) = "" Then MkDir p & "\assets\debug"
        CopyInto App.Path & "\assets\debug", p & "\assets\debug"
        CopyInto App.Path & "\framework", p
    Else
        Dim exeP As String
        exeP = """" & App.Path & "\Builder.exe" & """"
        Set WSHShell = CreateObject("WScript.Shell")
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\", "Create Emerald Project Here"
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon", exeP
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\", exeP & " ""%v"""
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\", "Create Emerald Project Here"
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon", exeP
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\", exeP & " ""%v"""
        
        MsgBox "Emerald Builder has been setuped on your computer .", 64
    End If
End Sub
Sub CopyInto(Src As String, Dst As String)
    Dim f As String
    f = Dir(Src & "\")
    Do While f <> ""
        FileCopy Src & "\" & f, Dst & "\" & f
        f = Dir()
    Loop
End Sub

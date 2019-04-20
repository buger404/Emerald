Attribute VB_Name = "Process"
Public Const Version As Long = 19042001
Public Function CheckFileName(name As String) As Boolean
    CheckFileName = ((InStr(name, "*") Or InStr(name, "\") Or InStr(name, "/") Or InStr(name, ":") Or InStr(name, "?") Or InStr(name, """") Or InStr(name, "<") Or InStr(name, ">") Or InStr(name, "|")) = 0)
End Function
Sub Main()
    If Command$ <> "" Then
        Dim appn As String, f As String, t As String, p As String
        Dim nList As String, xinfo As String, info() As String
        p = Replace(Command$, """", "")

        
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
                MsgBox "你的工程已经创建，我们已将最新的文件复制到你的文件夹中，你可以稍后引用它们。" & vbCrLf & vbCrLf & "注意：以下是更新Emerald后新增的文件，需要你手动引用（位于目录下的Core文件夹）：" & vbCrLf & nList, 64, "Emerald Builder"
                GoTo SkipName
            Else
                MsgBox "你的工程已经在使用最新的Emerald了。", 48, "Emerald Builder"
                Exit Sub
            End If
        End If

        appn = InputBox("输入你的工程名称", "Emerald Project")
        If CheckFileName(appn) = False Or appn = "" Then MsgBox "错误的工程名称。", 16, "Emerald Builder": Exit Sub
        
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
                If MsgBox("Emerald Builder 已经安装，你希望删除它吗？", vbYesNo + 48, "Emerald Builder") = vbYes Then
                    
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\"
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon"
                    
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\version"
                    
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\"
                    
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\"
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon"
                    
                    WSHShell.RegDelete "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\"
                    
                    MsgBox "Emerald Builder 已经从你的电脑上删除。", 48, "Emerald Builder"
                    
                    End
                Else
                    End
                End If
            Else
                If MsgBox("按下确定后更新你的 Emerald Builder .", 64 + vbYesNo, "Emerald Builder") = vbNo Then Exit Sub
            End If
        End If
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\", "在此处创建/更新Emerald工程"
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\icon", exeP
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\version", Version
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\shell\emerald\command\", exeP & " ""%v"""
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\", "在此处创建/更新Emerald工程"
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\icon", exeP
        
        WSHShell.RegWrite "HKEY_CLASSES_ROOT\Directory\Background\shell\emerald\command\", exeP & " ""%v"""
        
        MsgBox "Emerald Builder 成功安装在你的电脑上。", 64
        
FailOper:
        MsgBox "出了一些意外，无法完成部分操作。" & vbCrLf & Err.Description & "(" & Err.Number & ")", 48, "Emerald Builder"
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

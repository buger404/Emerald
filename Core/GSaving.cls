VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSaving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald 相关代码

Private Type SaveData
    Name As String
    value As String
    datacheck As String
End Type
Private Type SaveFile
    Data() As SaveData
End Type
Dim MySave As SaveFile
Public AutoSave As Boolean
Public sToken As Boolean, Path As String
Public Key As String
'========================================================
'   Init
    Public Function Create(Name As String, Optional BMKey) As Boolean
        Path = "C:\Users\" & PoolCreateObject("Wscript.Network").Username & "\AppData\Local\" & Name
        AutoSave = True
        
        If Not IsMissing(BMKey) Then Key = BMKey
        
        ReDim MySave.Data(0)
        If Dir(Path & "\", vbDirectory) = "" Then
            If Name <> "Emerald.Core" Then
                SysPage.index = 0: SysPage.Key = 0
                Dim Win As Object
                Set Win = StartScreenDialog(600, 400, SysPage)
                Set SysPage.Win = Win
                Do While Not Win.AcceptMark
                    DoEvents
                Loop
                Unload Win
                
                If SysPage.Key = 1 Then MkDir Path & "\": Create = True
            Else
                MkDir Path & "\"
            End If
        Else
            If Dir(Path & "\save") <> "" Then
                On Error GoTo ErrHandle
                Open Path & "\save" For Binary As #1
                Get #1, , MySave
                Close #1
                Call CheckSave
ErrHandle:
                If Err.Number <> 0 Then MsgBox "存档已经损坏！", 16, DisplayName: ReDim MySave.Data(0)
            End If
            Create = True
        End If
        
        sToken = Create
        
        If ESave Is Nothing Then Set ESave = Me
    End Function
'========================================================
'   RunTime
    Public Sub CheckSave()
        Dim msg As String
        For I = 1 To UBound(MySave.Data)
            If MySave.Data(I).value <> "" Then
                If MySave.Data(I).datacheck <> IIf(Key = "", BMEA(MySave.Data(I).value & I), BMEA(MySave.Data(I).value & I, Key)) Then
                    msg = msg & MySave.Data(I).Name & vbCrLf
                    MySave.Data(I).value = ""
                End If
            End If
        Next
        
        If msg <> "" Then
            Call Save
            MsgBox "已经发现并删除存档中以下非法数据：" & vbCrLf & msg, 64
        End If
    End Sub
'========================================================
'   Manager
    Public Function Count() As Long
        Count = UBound(MySave.Data)
    End Function
    Public Function GetName(index As Integer) As String
        GetName = MySave.Data(index).Name
    End Function
    Public Sub Save()
        If sToken = False Then Exit Sub
        Open Path & "\save" For Binary As #1
        Put #1, , MySave
        Close #1
    End Sub
    Public Function GetData(Name As String) As String
        For I = 1 To UBound(MySave.Data)
            If MySave.Data(I).Name = Name Then GetData = MySave.Data(I).value: Exit For
        Next
    End Function
    Public Sub PutData(Name As String, ByVal value As String)
        Dim fd As Boolean
        For I = 1 To UBound(MySave.Data)
            If MySave.Data(I).Name = Name Then
                MySave.Data(I).value = value
                MySave.Data(I).datacheck = IIf(Key = "", BMEA(value & I), BMEA(value & I, Key))
                fd = True
                Exit For
            End If
        Next
        If fd = False Then
            '新增
            ReDim Preserve MySave.Data(UBound(MySave.Data) + 1)
            MySave.Data(UBound(MySave.Data)).Name = Name
            MySave.Data(UBound(MySave.Data)).value = value
            MySave.Data(UBound(MySave.Data)).datacheck = IIf(Key = "", BMEA(value & UBound(MySave.Data)), BMEA(value & UBound(MySave.Data), Key))
        End If
        If AutoSave Then Call Save
    End Sub
'========================================================

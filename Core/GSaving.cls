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
    name As String
    value As String
    datacheck As String
End Type
Private Type SaveFile
    data() As SaveData
End Type
Dim MySave As SaveFile
Public AutoSave As Boolean
Public sToken As Boolean, path As String
Public Key As String
'========================================================
'   Init
    Public Function Create(name As String, Optional BMKey) As Boolean
        path = "C:\Users\" & CreateObject("Wscript.Network").Username & "\AppData\Local\" & name
        AutoSave = True
        
        If Not IsMissing(BMKey) Then Key = BMKey
        
        ReDim MySave.data(0)
        If Dir(path & "\", vbDirectory) = "" Then
            If name <> "Emerald.Core" Then
                SysPage.index = 0: SysPage.Key = 0
                Dim Win As Object
                Set Win = StartScreenDialog(600, 400, SysPage)
                Set SysPage.Win = Win
                Do While Not Win.AcceptMark
                    DoEvents
                Loop
                Unload Win
                
                If SysPage.Key = 1 Then MkDir path & "\"
            Else
                MkDir path & "\"
            End If
        Else
            If Dir(path & "\save") <> "" Then
                Open path & "\save" For Binary As #1
                Get #1, , MySave
                Close #1
                Call CheckSave
            End If
            Create = True
        End If
        
        sToken = Create
        
        Set ESave = Me
    End Function
'========================================================
'   RunTime
    Public Sub CheckSave()
        Dim msg As String
        For i = 1 To UBound(MySave.data)
            If MySave.data(i).value <> "" Then
                If MySave.data(i).datacheck <> IIf(Key = "", BMEA(MySave.data(i).value & i), BMEA(MySave.data(i).value & i, Key)) Then
                    msg = msg & MySave.data(i).name & vbCrLf
                    MySave.data(i).value = ""
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
        Count = UBound(MySave.data)
    End Function
    Public Function GetName(index As Integer) As String
        GetName = MySave.data(index).name
    End Function
    Public Sub Save()
        If sToken = False Then Exit Sub
        Open path & "\save" For Binary As #1
        Put #1, , MySave
        Close #1
    End Sub
    Public Function GetData(name As String) As String
        For i = 1 To UBound(MySave.data)
            If MySave.data(i).name = name Then GetData = MySave.data(i).value: Exit For
        Next
    End Function
    Public Sub PutData(name As String, ByVal value As String)
        Dim fd As Boolean
        For i = 1 To UBound(MySave.data)
            If MySave.data(i).name = name Then
                MySave.data(i).value = value
                MySave.data(i).datacheck = IIf(Key = "", BMEA(value & i), BMEA(value & i, Key))
                fd = True
                Exit For
            End If
        Next
        If fd = False Then
            '新增
            ReDim Preserve MySave.data(UBound(MySave.data) + 1)
            MySave.data(UBound(MySave.data)).name = name
            MySave.data(UBound(MySave.data)).value = value
            MySave.data(UBound(MySave.data)).datacheck = IIf(Key = "", BMEA(value & UBound(MySave.data)), BMEA(value & UBound(MySave.data), Key))
        End If
        If AutoSave Then Call Save
    End Sub
'========================================================

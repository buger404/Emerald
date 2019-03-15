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
Private Type SaveData
    name As String
    value As String
    datacheck As String
End Type
Private Type SaveFile
    Data() As SaveData
End Type
Dim MySave As SaveFile, path As String
Public sToken As Boolean
'========================================================
'   Init
    Public Function Create(name As String, displayname As String) As Boolean
        path = "C:\Users\" & CreateObject("Wscript.Network").Username & "\AppData\Local\" & name
        
        ReDim MySave.Data(0)
        If Dir(path & "\", vbDirectory) = "" Then
            Create = (MsgBox(displayname & "希望在你的电脑的以下位置储存数据 :" & vbCrLf & path & vbCrLf & "同意吗？（如果不同意该程序将永远保持在初始状态）", 48 Or vbYesNo, displayname) = vbYes)
            If Create Then MkDir path & "\"
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
    End Function
'========================================================
'   RunTime
    Public Sub CheckSave()
        Dim msg As String
        For i = 1 To UBound(MySave.Data)
            If MySave.Data(i).value <> "" Then
                If MySave.Data(i).datacheck <> BMEA(MySave.Data(i).value & i) Then
                    msg = msg & MySave.Data(i).name & vbCrLf
                    MySave.Data(i).value = ""
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
    Public Sub Save()
        If sToken = False Then Exit Sub
        Open path & "\save" For Binary As #1
        Put #1, , MySave
        Close #1
    End Sub
    Public Function GetData(name As String) As String
        For i = 1 To UBound(MySave.Data)
            If MySave.Data(i).name = name Then GetData = MySave.Data(i).value: Exit For
        Next
    End Function
    Public Sub PutData(name As String, ByVal value As String)
        Dim fd As Boolean
        For i = 1 To UBound(MySave.Data)
            If MySave.Data(i).name = name Then
                MySave.Data(i).value = value
                MySave.Data(i).datacheck = BMEA(value & i)
                fd = True
                Exit For
            End If
        Next
        If fd = False Then
            '新增
            ReDim Preserve MySave.Data(UBound(MySave.Data) + 1)
            MySave.Data(UBound(MySave.Data)).name = name
            MySave.Data(UBound(MySave.Data)).value = value
            MySave.Data(UBound(MySave.Data)).datacheck = BMEA(value & UBound(MySave.Data))
        End If
        Call Save
    End Sub
'========================================================

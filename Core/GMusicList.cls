VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMusicList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type MusicL
    name As String
    music As GMusic
End Type
Dim List() As MusicL
'========================================================
'   Init
    Private Sub Class_Initialize()
        ReDim List(0)
    End Sub
'========================================================
'   Manager
    Public Sub Create(ByVal folder As String)
        Dim file As String
        If Right(folder, 1) <> "\" Then folder = folder & "\"
        file = Dir(folder)
        
        Do While file <> ""
            ReDim Preserve List(UBound(List) + 1)
            Set List(UBound(List)).music = New GMusic
            List(UBound(List)).music.Create folder & file
            List(UBound(List)).name = file
            file = Dir()
        Loop
    End Sub
'========================================================
'   Play
    Public Sub Play(name As String)
        For i = 1 To UBound(List)
            If List(i).name = name Then List(i).music.Play: Exit For
        Next
    End Sub
'========================================================

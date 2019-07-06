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
'Emerald Ïà¹Ø´úÂë

Private Type MusicL
    name As String
    Music As GMusic
End Type
Public HotLoad As Boolean
Dim AssetsFolder() As String
Dim List() As MusicL
'========================================================
'   Init
    Private Sub Class_Initialize()
        ReDim List(0)
        ReDim AssetsFolder(0)
        Set EMusic = Me
    End Sub
'========================================================
'   Manager
    Public Sub Create(ByVal Folder As String)
        Dim File As String
        If Right(Folder, 1) <> "\" Then Folder = Folder & "\"
        File = Dir(Folder)
        
        If HotLoad Then
            ReDim Preserve AssetsFolder(UBound(AssetsFolder) + 1)
            AssetsFolder(UBound(AssetsFolder)) = Folder
            Exit Sub
        End If
        
        Do While File <> ""
            ReDim Preserve List(UBound(List) + 1)
            Set List(UBound(List)).Music = New GMusic
            List(UBound(List)).Music.Create Folder & File
            List(UBound(List)).name = File
            File = Dir()
        Loop
    End Sub
'========================================================
'   Play
    Public Sub Play(name As String)
        For i = 1 To UBound(List)
            If List(i).name = name Then List(i).Music.Play: Exit Sub
        Next
        
        If HotLoad Then
            For i = 1 To UBound(AssetsFolder)
                If IsExitAFile(AssetsFolder(i) & name) Then
                    ReDim Preserve List(UBound(List) + 1)
                    Set List(UBound(List)).Music = New GMusic
                    List(UBound(List)).Music.Create AssetsFolder(i) & name
                    List(UBound(List)).name = name
                    List(UBound(List)).Music.Play
                    Exit For
                End If
            Next
        End If
    End Sub
'========================================================

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Page
    p As GPage
    name As String
End Type
Private Type Transform
    kind As Integer
    time As Long
    during As Long
End Type
Dim Pages() As Page, Transform As Transform
Public ActivePage As String, CDC As Long, ODC As Long
'========================================================
'   Init
    Private Sub Class_Initialize()
        Set ECore = Me
        ReDim Pages(0)
        CDC = CreateCDC(GW, GH)
        ODC = CreateCDC(GW, GH)
    End Sub
    Public Sub Dispose()
        For i = 1 To UBound(Pages)
            Pages(i).p.Dispose
        Next
        DeleteObject CDC
        DeleteObject ODC
    End Sub
'========================================================
'   Manager
    Public Sub Add(p As GPage, name As String)
        ReDim Preserve Pages(UBound(Pages) + 1)
        With Pages(UBound(Pages))
            Set .p = p
            .name = name
        End With
    End Sub
'========================================================
'   Draw
    Public Sub Display()
        For i = 1 To UBound(Pages)
            If ActivePage = Pages(i).name Then Pages(i).p.Clear: Pages(i).p.Update: Pages(i).p.Display (CDC): Exit For
        Next
        If GetTickCount - Transform.time <= Transform.during Then Call Translation
        BitBlt GDC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
        If Mouse.State = 2 Then Mouse.State = 0
    End Sub
'========================================================
'   Transform
    Public Sub NewTransform(Optional kind As Integer = 0, Optional during As Long = 300)
        PaintDC CDC, ODC
        With Transform
            .during = during
            .kind = kind
            .time = GetTickCount
        End With
    End Sub
    Private Sub Translation()
        Dim pro As Single
        Select Case Transform.kind
            Case 0 'FadeIn
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, alpha:=1 - pro
            Case 1 'FadeOut
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 0, 0, 1)
                PaintDC ODC, CDC, alpha:=1 - pro
            Case 2 'ToRight
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * GW), alpha:=1
            Case 3 'ToLeft
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), alpha:=1
            Case 4 'ToUp
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), alpha:=1
            Case 5 'ToDown
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), alpha:=1
        End Select
    End Sub
'========================================================

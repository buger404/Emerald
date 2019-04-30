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
Dim mPage As String, aPage As Object
Dim OutOfWindow As Boolean
Dim FirstPlay As Boolean
Public CDC As Long, ODC As Long, GG As Long, pB As Long
'========================================================
'   Init
    Private Sub Class_Initialize()
        Set ECore = Me
        ReDim Pages(0)
        CDC = CreateCDC(GW, GH)
        ODC = CreateCDC(GW, GH)
        GdipCreateFromHDC CDC, GG
        GdipCreateSolidFill argb(255, 255, 255, 255), pB
    End Sub
    Public Sub Dispose()
        GdipDeleteGraphics GG
        GdipDeleteBrush pB
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
    Public Property Get ActivePage() As String
        ActivePage = mPage
    End Property
    Public Property Let ActivePage(nPage As String)
        For i = 1 To UBound(Pages)
            If mPage = Pages(i).name Then CallByName Pages(i).p.SrcClass, "Leave", VbMethod: Exit For
        Next
        mPage = nPage
        For i = 1 To UBound(Pages)
            If nPage = Pages(i).name Then
                CallByName Pages(i).p.SrcClass, "Enter", VbMethod
                Set aPage = Pages(i).p.SrcClass
                Exit For
            End If
        Next
    End Property
    Public Sub Wheel(Direction As Integer, Depth As Single)
        CallByName aPage, "Wheel", VbMethod, Direction, Depth
    End Sub
'========================================================
'   Draw
    Public Sub Display()
        Dim o As Boolean, t As Long
        
        t = Timer * 1000
        
        If ECore Is Me Then
            If (Not FirstPlay) Then
                FirstPlay = True: SysPage.OpenTime = GetTickCount: SysPage.Index = 1
                Call NewTransform
            End If
            If (Not SysPage.DoneMark) Then
                SysPage.Page.Clear
                SysPage.Page.Update
                SysPage.Page.Display CDC
                GoTo skipPage
            End If
        End If

        For i = 1 To UBound(Pages)
            If ActivePage = Pages(i).name Then
                Pages(i).p.Clear: Pages(i).p.Update: Pages(i).p.Display (CDC)
                o = (GetActiveWindow = 0)
                If o <> OutOfWindow Then
                    OutOfWindow = o
                    If OutOfWindow Then
                        CallByName Pages(i).p.SrcClass, "Lostfocus", VbMethod
                    Else
                        CallByName Pages(i).p.SrcClass, "Gotfocus", VbMethod
                    End If
                End If
                Exit For
            End If
        Next
        
skipPage:
        If GetTickCount - Transform.time <= Transform.during Then Call Translation
        BitBlt GDC, 0, 0, GW, GH, CDC, 0, 0, vbSrcCopy
        tFPS = tFPS + 1: FPSctt = FPSctt + (Timer * 1000 - t)
        
        If GetTickCount - FPSt >= 1000 Then
            FPSt = GetTickCount: FPS = tFPS: tFPS = 0
            FPSct = FPSctt: FPSctt = 0
        End If
        
        If Mouse.State = 2 Then Mouse.State = 0
    End Sub
'========================================================
'   Transform
    Public Sub NewTransform(Optional kind As Integer = 0, Optional during As Long = 300, Optional newPage)
        PaintDC CDC, ODC
        With Transform
            .during = during
            .kind = kind
            .time = GetTickCount
        End With
        
        If Not IsMissing(newPage) Then
            ActivePage = newPage
        End If
    End Sub
    Private Sub Translation()
        Dim pro As Single
        Select Case Transform.kind
            Case 0 'FadeIn
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, Alpha:=1 - pro
            Case 1 'FadeOut
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 0, 0, 1)
                PaintDC ODC, CDC, Alpha:=1 - pro
            Case 2 'ToRight
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * GW), Alpha:=1
            Case 3 'ToLeft
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), Alpha:=1
            Case 4 'ToUp
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), Alpha:=1
            Case 5 'ToDown
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), Alpha:=1
            Case 6 'ToRight(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, (pro * GW), Alpha:=1 - pro
            Case 7 'ToLeft(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, -(pro * GW), Alpha:=1 - pro
            Case 8 'ToUp(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, -(pro * GW), Alpha:=1 - pro
            Case 9 'ToDown(Fade)
                pro = Cubic((GetTickCount - Transform.time) / Transform.during, 0, 1, 1, 1)
                PaintDC ODC, CDC, 0, (pro * GW), Alpha:=1 - pro
            Case 10 'HighLight
                pro = (GetTickCount - Transform.time) / Transform.during
                Select Case pro
                    Case Is < 0.2
                        pro = Cubic(pro / 0.2, 0, 1, 1, 1)
                        GdipSetSolidFillColor pB, argb(pro * 255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 0.8
                        GdipSetSolidFillColor pB, argb(255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 1
                        pro = Cubic((pro - 0.8) / 0.2, 0, 0, 0, 1)
                        GdipSetSolidFillColor pB, argb((1 - pro) * 255, 255, 255, 255)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                End Select
            Case 11 'Dark
                pro = (GetTickCount - Transform.time) / Transform.during
                Select Case pro
                    Case Is < 0.2
                        pro = Cubic(pro / 0.2, 0, 1, 1, 1)
                        GdipSetSolidFillColor pB, argb(pro * 255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 0.8
                        GdipSetSolidFillColor pB, argb(255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                    Case Is < 1
                        pro = Cubic((pro - 0.8) / 0.2, 0, 0, 0, 1)
                        GdipSetSolidFillColor pB, argb((1 - pro) * 255, 0, 0, 0)
                        GdipFillRectangle GG, pB, 0, 0, GW, GH
                End Select
        End Select
    End Sub
'========================================================

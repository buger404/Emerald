VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
    Dim Page As GPage
    Public ErrNumber As Long, ErrDes As String, ErrPage As String
    Public ODC As Long
'======================================

Public Sub Update()
    'Error Screen
    'All the errors will be catched here .
    
    Page.Clear
    PaintDC ODC, Page.CDC, 0, 0, 0, 0, GW, GH, 1
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(180, 0, 0, 0)
    
    Page.Writes "Ouch , an error occurred .", GW / 2, GH / 2 - 80, 24, argb(255, 254, 84, 57), style:=FontStyleBold, align:=StringAlignmentCenter
    Page.Writes "# " & ErrNumber & " #", GW / 2, GH / 2 - 40, 24, argb(255, 254, 84, 57), style:=FontStyleBold, align:=StringAlignmentCenter
    Page.Writes ErrDes & " ...", GW / 2, GH / 2, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes "Page name : " & ErrPage, GW / 2, GH / 2 + 50, 18, argb(255, 255, 255, 255), align:=StringAlignmentCenter
    Page.Writes "You can feedback this error .", GW / 2, GH / 2 + 150, 18, argb(160, 255, 255, 255), align:=StringAlignmentCenter
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    'Mouse Wheel
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    'Animation Message Receive
End Sub
Public Sub Enter()
    'Page Enter
End Sub
Public Sub Leave()
    'Page Leave
End Sub
Public Sub Gotfocus()
    'Window Get Focus
End Sub
Public Sub Lostfocus()
    'Window Lose Focus
End Sub
Public Sub AnimationDone(id As Integer)
    'Animation Done
    'id£ºFinishing Anmation id
End Sub
Private Sub Class_Initialize()
    'Create Page
    Set Page = New GPage
    Page.Create Me
    'Create
    ECore.Add Page, "ErrorScreen"
    ODC = CreateCDC(GW, GH)
End Sub

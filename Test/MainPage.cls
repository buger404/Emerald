VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MainPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   页面绘制器
    Dim Page As GPage, Music As GMusic
    Dim test As Boolean, pos As Single
'======================================

Public Sub Update()
    '绘制事件
    Dim Map() As Single, MapTotal As Long, w As Long, h As Long
    
    Page.Clear argb(170, 255, 255, 255)
    
    On Error Resume Next
    Map = Music.GetMusicMap
    For i = 0 To UBound(Map)
        MapTotal = MapTotal + Map(i)
    Next
    
    Dim pro As Single
    Page.Paint 0, 0, GH / 2 - 1, GW, 4, argb(255, 222, 222, 222)
    
    Page.Paint 1, GW / 2 - 256 / 2, GH / 2 - 256 / 2, 256, 256, argb(255, 222, 222, 222)
    pro = MapTotal / 100000
    If pro > 1 Then pro = 1
    Page.PaintArc GW / 2 - 256 / 2, GH / 2 - 256 / 2, 256, 256, pro * 360, Color:=argb(255, 0, 176, 240), size:=3, style:=1
    Page.Paint 1, GW / 2 - 256 / 2, GH / 2 - 256 / 2, 256, 256, argb(255 * pro, 0, 176, 240)

    For i = 0 To 20
        pro = Map(i) / 3000
        If pro > 1 Then pro = 1
        Page.PaintLine GW / 2 + Int(GW / 30) * i + 148, GH / 2 - GH * pro / 2, GW / 2 + Int(GW / 30) * i + 148, GH / 2 + GH * pro / 2, argb(255, 0, 176, 240), 4
        Page.PaintLine GW / 2 - Int(GW / 30) * i - 148, GH / 2 - GH * pro / 2, GW / 2 - Int(GW / 30) * i - 148, GH / 2 + GH * pro / 2, argb(255, 0, 176, 240), 4
    Next
    Page.DrawImage "music.png", GW / 2, GH / 2, Alpha:=1, pos:=1
    
    Dim ret As Integer
    ret = Page.ShowHScrollBar(pos, 20, GH - 40, GW - 40, 20, argb(255, 222, 222, 222), argb(255, 0, 183, 195))
    If ret <= 1 Then pos = Music.Position / Music.length
    If ret = 3 Then
        Music.Position = pos * Music.length
    End If
    Page.Writes ToTime(pos * Music.length) & " / " & ToTime(Music.length), 20, GH - 65, 16, argb(255, 27, 27, 27)
    
    If Not Music.IsPlaying Then Music.Play
End Sub
Public Sub PlayNew(Path As String)
    Music.Dispose
    Music.Create Path
    Music.Play
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub Enter()
    '页面进入事件
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()
    '创建页面绘制器
    Set Page = New GPage
    Page.Create Me
    '导入游戏资源
    Page.Res.NewImages App.Path & "\assets"
    Page.Res.NewImages App.Path & "\assets\debug"
    '创建页面
    ECore.Add Page, "MainPage"
    
    Set Music = New GMusic
    PlayNew App.Path & "\music\Anticipation.mp3"
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub

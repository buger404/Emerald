VERSION 5.00
Begin VB.Form Debuginfo 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Emerald Console"
   ClientHeight    =   7596
   ClientLeft      =   0
   ClientTop       =   -12
   ClientWidth     =   12240
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.2
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000005&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   633
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1020
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer UpdateTimer 
      Interval        =   10
      Left            =   216
      Top             =   240
   End
End
Attribute VB_Name = "Debuginfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Emerald 相关代码
Dim Page As GPage, Console As GDebug, sh As New aShadow
Dim WDC As Long
Dim ScrollMode As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Not Console.InputAllow Then Exit Sub
    
    If KeyAscii = 13 Then
        Console.ApplyCmd
    ElseIf KeyAscii = vbKeyBack Then
        If Len(Console.InputingText) > 0 Then Console.InputingText = Left(Console.InputingText, Len(Console.InputingText) - 1)
    ElseIf KeyAscii >= 0 And KeyAscii <= 26 Then
        On Error Resume Next
        If KeyAscii = 3 Then
            Clipboard.Clear
            Clipboard.SetText Console.InputingText
        End If
        If KeyAscii = 22 Then
            Console.InputingText = Console.InputingText & Clipboard.GetText
        End If
        If KeyAscii = 24 Then
            Clipboard.Clear
            Clipboard.SetText Console.InputingText
            Console.InputingText = ""
        End If
    Else
        Console.InputingText = Console.InputingText & Chr(KeyAscii)
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not Console.InputAllow Then Exit Sub
    
    Console.ApplyKey KeyCode
End Sub

Private Sub Form_Load()
    Set Page = New GPage
    Set Console = New GDebug
    
    Page.Create Console
    Page.Res.NewImages App.path & "\assets\debug", 64, 64
    
    Set Console.Page = Page
    Console.PageMark = 1
    
    Me.Width = 1020 * Screen.TwipsPerPixelX: Me.Height = 633 * Screen.TwipsPerPixelY
    Console.GW = 1020: Console.GH = 633
    Console.InitConsole
    
    WDC = CreateCDC(Console.GW, Console.GH)
    DeleteObject Page.CDC
    Page.CDC = WDC
    Dim g As Long
    GdipCreateFromHDC WDC, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipSetTextRenderingHint g, TextRenderingHintAntiAlias
    GdipDeleteGraphics Page.GG
    Page.GG = g
    
    SetWindowLongA Me.Hwnd, GWL_EXSTYLE, GetWindowLongA(Me.Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    BlurWindow Me.Hwnd
    
    With sh
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 12
            .Transparency = 32
        End If
    End With

End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
    If y <= 40 And button = 1 Then
        ReleaseCapture
        SendMessageA Me.Hwnd, WM_SYSCOMMAND, SC_MOVE Or HTCAPTION, 0
    End If
    Call Form_MouseMove(button, Shift, x, y)
    If button = 1 Then
        If Console.NeedScroll And x >= Console.GW - 20 And x <= Console.GW - 10 And y >= 40 Then
            SetCapture Me.Hwnd
            ScrollMode = True
        End If
    End If
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
    If button = 1 Then
        If Console.NeedScroll And ScrollMode Then
            Dim MaxY As Single
            MaxY = (Console.CuY - Console.GH + 80) / 3220
            Console.SY = (y - 60) / (Console.GH - 60 - 20) * MaxY
            If Console.SY < 0 Then Console.SY = 0
            If Console.SY > MaxY Then Console.SY = MaxY
        End If
    End If
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseMove(button, Shift, x, y)
    If ScrollMode Then ReleaseCapture: ScrollMode = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Page.Dispose
    Set sh = Nothing
End Sub

Public Sub UpdateTimer_Timer()
    If EmeraldInstalled = False Then Exit Sub
    If Me.Visible = False Then Exit Sub
    
    'If GetActiveWindow <> Me.Hwnd Then Exit Sub
    
    Page.Update

    Dim bs As BLENDFUNCTION, sz As size
    Dim SrcPoint As POINTAPI
    With bs
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    sz.cx = Console.GW: sz.cy = Console.GH

    UpdateLayeredWindow Me.Hwnd, WDC, ByVal 0&, sz, WDC, SrcPoint, 0, bs, &H2
End Sub

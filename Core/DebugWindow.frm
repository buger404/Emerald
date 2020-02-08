VERSION 5.00
Begin VB.Form DebugWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4920
   ForeColor       =   &H008C8080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   55
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer UpdateTimer 
      Interval        =   20
      Left            =   6600
      Top             =   120
   End
   Begin VB.Label touchArea 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   516
      Index           =   0
      Left            =   3144
      TabIndex        =   0
      Top             =   72
      Visible         =   0   'False
      Width           =   516
   End
End
Attribute VB_Name = "DebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Emerald 相关代码
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Dim Page As GPage, Charge As GDebug
Dim WDC As Long
Private Sub Form_Load()
    Set Page = New GPage
    Set Charge = New GDebug
    
    Page.IsSystem = True
    Page.Create Charge
    Page.Res.NewImages App.path & "\assets\debug", 48, 48
    If Dir(App.path & "\assets\sets\profile.png") <> "" Then
        Page.Res.newImage App.path & "\assets\sets\profile.png", 36, 36, "profile.png"
    Else
        Page.Res.newImage App.path & "\assets\debug\icon.png", 36, 36, "profile.png"
    End If
    Page.Res.ClipCircle "profile.png"
    
    Set Charge.Page = Page
    
    Me.Width = 410 * Screen.TwipsPerPixelX: Me.Height = 55 * Screen.TwipsPerPixelY
    Charge.GW = Me.ScaleWidth: Charge.GH = Me.ScaleHeight
    
    WDC = CreateCDC(Charge.GW, Charge.GH)
    DeleteDC Page.CDC
    Page.CDC = WDC
    Dim g As Long
    PoolCreateFromHdc WDC, g
    GdipSetSmoothingMode g, SmoothingModeAntiAlias
    GdipSetTextRenderingHint g, TextRenderingHintAntiAlias
    PoolDeleteGraphics Page.GG
    Page.GG = g
    
    Me.Move Screen.Width / 2 - Me.ScaleWidth * Screen.TwipsPerPixelX / 2, Screen.Height - GetTaskbarHeight - 55 * Screen.TwipsPerPixelY
    
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    For I = 1 To 5
        Load touchArea(I)
        With touchArea(I)
            .Visible = True
            .ZOrder
            .Move Me.ScaleWidth - 48 * I, 54 / 2 - 48 / 2, 48, 48
            Select Case I
                Case 3
                    .ToolTipText = "控制台"
                Case 1
                    .ToolTipText = "鼠标状态指示&点击检测矩形"
                Case 5
                    .ToolTipText = "界面设计器吸附模式"
                Case 4
                    .ToolTipText = "显示/不显示绘制矩形和坐标"
                Case 2
                    .ToolTipText = "存档数据管理"
            End Select
        End With
    Next
    
    SetWindowLongA Me.Hwnd, GWL_EXSTYLE, GetWindowLongA(Me.Hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    BlurWindow Me.Hwnd
End Sub
Public Function GetTaskbarHeight() As Long
    Dim lRes As Long
    Dim rectVal As RECT
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = (((Screen.Height / Screen.TwipsPerPixelY) - rectVal.Bottom) * Screen.TwipsPerPixelY)
End Function
Private Sub Form_Unload(Cancel As Integer)
    Page.Dispose
End Sub

Public Sub touchArea_Click(index As Integer)
    Select Case index
        Case 3
            Debuginfo.Visible = Not Debuginfo.Visible
        Case 1
            Debug_mouse = Not Debug_mouse
        Case 5
            Debug_umode = Debug_umode + 1
            If Debug_umode > 2 Then Debug_umode = 0
            Select Case Debug_umode
                Case 0: touchArea(5).ToolTipText = "界面设计器吸附模式"
                Case 1: touchArea(5).ToolTipText = "物件吸附模式"
                Case 2: touchArea(5).ToolTipText = "网格吸附模式"
            End Select
        Case 4
            Debug_focus = Not Debug_focus
            Debug_pos = Not Debug_pos
        Case 2
            If Not Debug_data Then
                SysPage.DoneMark = False: SysPage.DoneStep = 0
                SysPage.OpenTime = GetTickCount: SysPage.index = 3
                Call ECore.NewTransform
                Debug_data = True
            Else
                SysPage.DoneMark = True
                Call ECore.NewTransform
                Debug_data = False
            End If
    End Select
End Sub

Private Sub UpdateTimer_Timer()
    If EmeraldInstalled = False Then Exit Sub
    Page.Clear argb(200, 32, 32, 39)
    Page.Update
    
    Dim bs As BLENDFUNCTION, sz As size
    Dim SrcPoint As POINTAPI
    With bs
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    sz.cx = 410: sz.cy = 55

    UpdateLayeredWindow Me.Hwnd, Page.CDC, ByVal 0&, sz, Page.CDC, SrcPoint, 0, bs, &H2
End Sub

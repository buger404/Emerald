VERSION 5.00
Begin VB.Form DebugWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   936
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7032
   ForeColor       =   &H008C8080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   78
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   586
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
      Height          =   730
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   730
   End
End
Attribute VB_Name = "DebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Emerald 相关代码
Dim Page As GPage, Charge As GDebug
Dim WDC As Long
Private Sub Form_Load()
    Set Page = New GPage
    Set Charge = New GDebug
    
    Page.Create Charge
    Page.Res.NewImages App.path & "\assets\debug", 64, 64
    
    Set Charge.Page = Page
    
    Me.Width = 586 * Screen.TwipsPerPixelX: Me.Height = 78 * Screen.TwipsPerPixelY
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
    
    Me.Move Screen.Width / 2 - Me.ScaleWidth * Screen.TwipsPerPixelX / 2, 0
    
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    For I = 1 To 5
        Load touchArea(I)
        With touchArea(I)
            .Visible = True
            .ZOrder
            .Move Me.ScaleWidth - 10 - 64 * I, 78 / 2 - 64 / 2, 64, 64
            Select Case I
                Case 1
                    .ToolTipText = "详细信息窗口"
                Case 2
                    .ToolTipText = "鼠标状态指示&点击检测矩形"
                Case 3
                    .ToolTipText = "显示/不显示绘制矩形"
                Case 4
                    .ToolTipText = "显示/不显示绘制源坐标"
                Case 5
                    .ToolTipText = "存档数据管理"
            End Select
        End With
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Page.Dispose
End Sub

Public Sub touchArea_Click(index As Integer)
    Select Case index
        Case 1
            Debuginfo.Visible = Not Debuginfo.Visible
        Case 2
            Debug_mouse = Not Debug_mouse
        Case 3
            Debug_focus = Not Debug_focus
        Case 4
            Debug_pos = Not Debug_pos
        Case 5
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
    Page.Clear
    Page.Update
    Page.Display Me.hdc
End Sub

VERSION 5.00
Begin VB.Form DebugWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7030
   LinkTopic       =   "Form1"
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer UpdateTimer 
      Interval        =   1000
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
Dim Page As GPage, Charge As GDebug, Sh As New aShadow
Private Sub Form_Load()
    Set Page = New GPage
    Set Charge = New GDebug
    
    Page.Create Charge
    Page.NewImages App.path & "\assets\debug", 80, 80
    
    Set Charge.Page = Page
    
    Charge.GW = Me.ScaleWidth: Charge.GH = Me.ScaleHeight
    
    With Sh
        If .Shadow(Me) Then
            .Color = RGB(0, 0, 0)
            .Depth = 12
            .Transparency = 18
        End If
    End With
    
    Me.Move Screen.Width / 2 - Me.ScaleWidth * Screen.TwipsPerPixelX / 2, 0
    
    SetWindowPos Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    For i = 1 To 3
        Load touchArea(i)
        With touchArea(i)
            .Visible = True
            .ZOrder
            .Move Me.ScaleWidth - 10 - 80 * i, 5, 80, 80
        End With
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Sh = Nothing
End Sub

Private Sub touchArea_Click(Index As Integer)
    Select Case Index
        Case 1
            Debuginfo.Show
        Case 3
            Debug_focus = Not Debug_focus
    End Select
End Sub

Private Sub UpdateTimer_Timer()
    Page.Clear
    Page.Update
    Page.Display Me.hdc
End Sub

VERSION 5.00
Begin VB.Form EmeraldWindow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00212121&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald 授权中心"
   ClientHeight    =   4068
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7920
   ControlBox      =   0   'False
   ForeColor       =   &H00FF3E64&
   Icon            =   "EmeraldWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox qIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   924
      Left            =   360
      Picture         =   "EmeraldWindow.frx":1BCC2
      ScaleHeight     =   924
      ScaleWidth      =   924
      TabIndex        =   0
      Top             =   984
      Width           =   924
   End
   Begin VB.Label Back 
      BackColor       =   &H00FFFFFF&
      Height          =   4164
      Left            =   -72
      TabIndex        =   5
      Top             =   -48
      Width           =   1716
   End
   Begin VB.Label AcBtn 
      Alignment       =   2  'Center
      BackColor       =   &H00CC7A00&
      Caption         =   "接受"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   324
      Left            =   5184
      TabIndex        =   4
      Top             =   3456
      Width           =   1032
   End
   Begin VB.Label ReBtn 
      Alignment       =   2  'Center
      BackColor       =   &H007C7C7C&
      Caption         =   "拒绝"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   324
      Left            =   6504
      TabIndex        =   3
      Top             =   3456
      Width           =   1032
   End
   Begin VB.Label Content 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "该应用要求在您的计算机的下列位置储存文件："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   276
      Left            =   2040
      TabIndex        =   2
      Top             =   816
      Width           =   4284
   End
   Begin VB.Label Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "读写存档"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CC7A00&
      Height          =   324
      Left            =   2040
      TabIndex        =   1
      Top             =   384
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00CC7A00&
      X1              =   172
      X2              =   630
      Y1              =   266
      Y2              =   266
   End
End
Attribute VB_Name = "EmeraldWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pkey As Integer
Public Function NewPermissionDialog(nTitle As String, nContent As String) As Integer
    Title.Caption = nTitle
    Content.Caption = nContent
    pkey = -1
    Me.Show
    Do While pkey = -1
        Sleep 32: DoEvents
    Loop
    Me.Hide
    NewPermissionDialog = pkey
    Unload Me
End Function

Private Sub AcBtn_Click()
    pkey = 1
End Sub

Private Sub Form_Load()
    qIcon.top = Me.Height / Screen.TwipsPerPixelY / 2 - qIcon.Height / 2 - 30
End Sub

Private Sub ReBtn_Click()
    pkey = 0
End Sub

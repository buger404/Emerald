VERSION 5.00
Begin VB.Form MainWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald Builder"
   ClientHeight    =   4320
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   7524
   Icon            =   "MainWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   5616
      ScaleHeight     =   1440
      ScaleWidth      =   1440
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1440
   End
   Begin Emerald_Builder.EEdit InputBox 
      Height          =   300
      Left            =   432
      TabIndex        =   3
      Top             =   3048
      Visible         =   0   'False
      Width           =   6828
      _ExtentX        =   12044
      _ExtentY        =   529
      Content         =   "EEdit1"
      ForeColor       =   8422784
      BorderColor     =   13556506
      Alignment       =   0
      LockInput       =   0   'False
   End
   Begin Emerald_Builder.EButton Buttons 
      Height          =   420
      Index           =   0
      Left            =   6312
      TabIndex        =   2
      Top             =   3576
      Visible         =   0   'False
      Width           =   948
      _ExtentX        =   1672
      _ExtentY        =   741
      DefaultColor    =   15592941
      HoverColor      =   13556250
      ForeColor       =   8422784
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Content         =   "OK"
      Align           =   0
   End
   Begin VB.Label Content 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Content"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808580&
      Height          =   276
      Left            =   480
      TabIndex        =   1
      Top             =   768
      Width           =   768
   End
   Begin VB.Label Title 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   16.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CEDB1A&
      Height          =   420
      Left            =   456
      TabIndex        =   0
      Top             =   336
      Width           =   648
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Emerald Ïà¹Ø´úÂë

Public Key As Integer
Public Sub NewDialog(t As String, c As String, Pic As String, InputMode As Boolean, b())
    Key = 0
    PicBox.Visible = (Pic <> "")
    
    For i = 1 To Buttons.UBound
        Unload Buttons(i)
    Next
    
    Title.Caption = t
    Content.Caption = c
    
    If Pic <> "" Then PicBox.Picture = LoadPicture(Pic)
    
    InputBox.Visible = InputMode
    
    For i = 0 To UBound(b)
        Load Buttons(Buttons.UBound + 1)
        With Buttons(Buttons.UBound)
            .Content = b(i)
            .Top = Buttons(0).Top
            .Left = Me.ScaleWidth - (20 + Buttons(0).Width) * (UBound(b) - i + 1) - 10
            .Width = Buttons(0).Width
            .Height = Buttons(0).Height
            .Visible = True
        End With
    Next
    
    InputBox.Content = ""
End Sub

Private Sub Buttons_Click(Index As Integer)
    Key = Index
    Me.Hide
End Sub


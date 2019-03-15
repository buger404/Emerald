VERSION 5.00
Begin VB.Form TestWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emerald.Test"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   644
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer DrawTimer 
      Interval        =   10
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "TestWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GTest As GTest, EC As GMan

Private Sub DrawTimer_Timer()
    EC.Display
End Sub

Private Sub Form_Load()
    StartEmerald Me.Hwnd, Me.ScaleWidth, Me.ScaleHeight
    MakeFont "Î¢ÈíÑÅºÚ"
    
    Set EC = New GMan
    Set GTest = New GTest
    
    EC.ActivePage = "TestPage"
    DrawTimer.Enabled = True
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateMouse X, Y, 1, button
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
    If Mouse.state = 0 Then UpdateMouse X, Y, 0, button
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
    UpdateMouse X, Y, 2, button
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DrawTimer.Enabled = False
    EndEmerald
End Sub

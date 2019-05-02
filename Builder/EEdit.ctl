VERSION 5.00
Begin VB.UserControl EEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   2880
   ScaleWidth      =   3840
   Begin VB.Timer Ani 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3048
      Top             =   360
   End
   Begin VB.TextBox mEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.2
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1044
      Left            =   936
      TabIndex        =   0
      Text            =   "..."
      Top             =   960
      Width           =   2172
   End
   Begin VB.Line aline 
      BorderColor     =   &H00CEDB1A&
      X1              =   888
      X2              =   2496
      Y1              =   2592
      Y2              =   2592
   End
End
Attribute VB_Name = "EEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald 相关代码

Event Change()
Event Commit()
Dim dic As Integer, mlText As String
Public Property Get LockInput() As Boolean
    LockInput = mEdit.Locked
End Property
Public Property Let LockInput(b As Boolean)
    mEdit.Locked = b
End Property
Public Property Get Alignment() As AlignmentConstants
    Alignment = mEdit.Alignment
End Property
Public Property Let Alignment(a As AlignmentConstants)
    mEdit.Alignment = a
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mEdit.ForeColor
End Property
Public Property Let ForeColor(c As OLE_COLOR)
    mEdit.ForeColor = c
End Property
Public Property Get Backcolor() As OLE_COLOR
    Backcolor = mEdit.Backcolor
End Property
Public Property Let Backcolor(c As OLE_COLOR)
    mEdit.Backcolor = c
    UserControl.Backcolor = c
End Property
Public Property Get Bordercolor() As OLE_COLOR
    Bordercolor = aline.Bordercolor
End Property
Public Property Let Bordercolor(c As OLE_COLOR)
    aline.Bordercolor = c
End Property
Public Property Get Font() As StdFont
    Set Font = mEdit.Font
End Property
Public Property Set Font(f As StdFont)
    Set mEdit.Font = f
End Property
Public Property Get Content() As String
    Content = mEdit.Text
End Property
Public Property Let Content(c As String)
    mEdit.Text = c
End Property

Private Sub Ani_Timer()
    Dim mw As Long, rv As Long, em As Boolean
    mw = UserControl.Width / 15
    
    rv = aline.X1 + IIf(dic = 1, -mw, mw)
    If aline.X2 < rv Then em = True: aline.Visible = False
    If rv < 0 Then em = True: aline.X1 = 0: aline.X2 = UserControl.Width
        
    If em Then Ani.Enabled = False: Exit Sub
    
    aline.X1 = rv
    aline.X2 = aline.X2 + IIf(dic = 1, mw, -mw)
End Sub
Private Sub mEdit_Change()
    RaiseEvent Change
End Sub
Private Sub mEdit_GotFocus()
    mlText = mEdit.Text
    mEdit.Backcolor = GetFocusColor
    dic = 0: Ani.Enabled = True
End Sub
Private Sub mEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If mEdit.Text <> mlText Then mlText = mEdit.Text: RaiseEvent Commit
    End If
End Sub
Private Sub mEdit_LostFocus()
    If mEdit.Text <> mlText Then mlText = mEdit.Text: RaiseEvent Commit
    mEdit.Backcolor = UserControl.Backcolor
    aline.Visible = True
    dic = 1: Ani.Enabled = True
End Sub
Private Function GetFocusColor() As Long
    Dim c(3) As Byte, c2(3) As Byte, r(3) As Long, r2(3) As Long
    CopyMemory c(0), UserControl.Backcolor, 4
    CopyMemory c2(0), mEdit.ForeColor, 4
    
    For i = 0 To 3: r(i) = c(i): Next
    For i = 0 To 3: r2(i) = c2(i): Next
    
    '手动混合颜色
    For i = 0 To 3: r(i) = r(i) * 0.8 + r2(i) * 0.2: Next
    
    GetFocusColor = RGB(r(0), r(1), r(2))
End Function

Private Sub UserControl_InitProperties()
    mEdit.Backcolor = RGB(255, 255, 255)
    mEdit.ForeColor = RGB(192, 192, 192)
    
    mEdit.Text = Extender.name
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mEdit.Backcolor = Val(PropBag.ReadProperty("BackColor", RGB(255, 255, 255)))
    mEdit.Text = PropBag.ReadProperty("Content", Extender.name)
    mEdit.ForeColor = PropBag.ReadProperty("ForeColor", RGB(192, 192, 192))
    aline.Bordercolor = PropBag.ReadProperty("BorderColor", aline.Bordercolor)
    mEdit.Alignment = PropBag.ReadProperty("Alignment", AlignmentConstants.vbLeftJustify)
    mEdit.Locked = PropBag.ReadProperty("LockInput", False)
    
    UserControl.Backcolor = mEdit.Backcolor
    
    Set mEdit.Font = PropBag.ReadProperty("Font", mEdit.Font)
    
    Call UserControl_Resize
    
    mlText = mEdit.Text
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    mEdit.Move 0, 0, UserControl.Width, UserControl.Height - Screen.TwipsPerPixelY
    aline.X1 = 0: aline.Y1 = UserControl.Height - Screen.TwipsPerPixelY
    aline.X2 = UserControl.Width: aline.Y2 = aline.Y1
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", mEdit.Backcolor, RGB(255, 255, 255)
    PropBag.WriteProperty "Content", mEdit.Text, Extender.name
    PropBag.WriteProperty "ForeColor", mEdit.ForeColor
    PropBag.WriteProperty "BorderColor", aline.Bordercolor
    PropBag.WriteProperty "Alignment", mEdit.Alignment
    PropBag.WriteProperty "LockInput", mEdit.Locked
    
    PropBag.WriteProperty "Font", mEdit.Font, mEdit.Font
End Sub


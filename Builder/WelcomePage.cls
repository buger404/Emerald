VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WelcomePage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   ҳ�������
    Public Page As GPage
'======================================

Public Sub Update()
    '�����¼�
    
    If InstalledPath <> "" Then GoTo OtherPage
    
    Dim BtnStr As String
    
    Page.Clear argb(255 * 0.85, 242, 242, 249) '��ջ���
    
    If Repaired Then
        Page.Writes "Something Wrong", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
        Page.Writes "There is something wrong with your Builder .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
        BtnStr = "Repair"
        GoTo Draw
    End If
    
    If InstalledPath <> "" Then
        If (Not IsUpdate) Then
            Page.Writes "Say Goodbye to Emerald Builder", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
            Page.Writes "Are you really going to uninstall me ?", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
            BtnStr = "Uninstall"
        Else
            Page.Writes "New Emerald Builder!", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
            Page.Writes "Make everything better .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
            BtnStr = "Update"
        End If
    End If
    
    If InstalledPath = "" Then
        Page.Writes "Welcome to Emerald Builder", GW / 2, 200, Color:=argb(255, 27, 27, 27), size:=36, align:=StringAlignmentCenter, animation:=1
        Page.Writes "Let's create our own games together .", GW / 2, 260, Color:=argb(150, 27, 27, 27), size:=18, align:=StringAlignmentCenter, animation:=2
        BtnStr = "Setup"
    End If
        
Draw:
    If Page.ShowColorButton(0, GW / 2 - 130 / 2, GH - 150, 130, 40, BtnStr, argb(255, 255, 255, 255), argb(255, 85, 219, 211), argb(255, 55, 189, 181), size:=16) = 3 Then
        ECore.NewTransform transFadeIn, 700, "SetupPage"
        CmdMark = BtnStr
        If BtnStr = "Uninstall" Then
            Call Uninstall
        Else
            Call Setup
        End If
    End If

    Exit Sub
    
OtherPage:
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    '����ҳ�������
    Set Page = New GPage
    Page.Create Me
    '������Ϸ��Դ
    Page.Res.NewImages App.Path & "\assets"
    '����ҳ��
    ECore.Add Page, "WelcomePage"
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
        Page.NewAnimation False, "FadeIn", 2, 500
        Page.NewAnimation False, "FadeIn", 2, 500
    '===============================================
End Sub
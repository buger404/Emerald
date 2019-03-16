VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================================
    'Animation Progress
    Public Function GetProgress_0(pro As Single) As Single
        GetProgress_0 = pro
    End Function
    Public Function GetProgress_1(pro As Single) As Single
        GetProgress_1 = Cubic(pro, 0, 0, 1, 1)
    End Function
    Public Function GetProgress_2(pro As Single) As Single
        GetProgress_2 = Cubic(pro, 0, 0, 0, 1)
    End Function
    Public Function GetProgress_3(pro As Single) As Single
        GetProgress_3 = Cubic(pro, 0, 1, 1, 1)
    End Function
'================================================================
    'Animation Project
    Public Sub FadeIn(x, y, w, h, alpha, pro As Single)
        alpha = alpha * pro
    End Sub
    Public Sub FadeOut(x, y, w, h, alpha, pro As Single)
        alpha = alpha - alpha * pro
    End Sub
    Public Sub FlyInFromLeft(x, y, w, h, alpha, pro As Single)
        x = x * pro
    End Sub
    Public Sub FlyInFromRight(x, y, w, h, alpha, pro As Single)
        x = x + (GW - x) * pro
    End Sub
    Public Sub FlyInFromTop(x, y, w, h, alpha, pro As Single)
        y = y * pro
    End Sub
    Public Sub FlyInFromBottom(x, y, w, h, alpha, pro As Single)
        y = y + (GH - y) * pro
    End Sub
    Public Sub FlyInFromLeftFade(x, y, w, h, alpha, pro As Single)
        x = x * pro
        alpha = alpha * pro
    End Sub
    Public Sub FlyInFromRightFade(x, y, w, h, alpha, pro As Single)
        x = x + (GW - x) * pro
        alpha = alpha * pro
    End Sub
    Public Sub FlyInFromTopFade(x, y, w, h, alpha, pro As Single)
        y = y * pro
        alpha = alpha * pro
    End Sub
    Public Sub FlyInFromBottomFade(x, y, w, h, alpha, pro As Single)
        y = y + (GH - y) * pro
        alpha = alpha * pro
    End Sub
'================================================================

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
'Emerald Ïà¹Ø´úÂë

'================================================================
    'Animation Progress
    Public Function GetProgress_0(pro As Single) As Single
        If pro < 0 Then GetProgress_0 = 0: Exit Function
        GetProgress_0 = pro
    End Function
    Public Function GetProgress_1(pro As Single) As Single
        If pro < 0 Then GetProgress_1 = 0: Exit Function
        GetProgress_1 = Cubic(pro, 0, 0, 1, 1)
    End Function
    Public Function GetProgress_2(pro As Single) As Single
        If pro < 0 Then GetProgress_2 = 0: Exit Function
        GetProgress_2 = Cubic(pro, 0, 0, 0, 1)
    End Function
    Public Function GetProgress_3(pro As Single) As Single
        If pro < 0 Then GetProgress_3 = 0: Exit Function
        GetProgress_3 = Cubic(pro, 0, 1, 1, 1)
    End Function
'================================================================
    'Animation Solution
    Public Sub FadeIn(X, y, w, h, alpha, pro As Single)
        alpha = alpha * pro
    End Sub
    Public Sub FadeOut(X, y, w, h, alpha, pro As Single)
        alpha = alpha - alpha * pro
    End Sub
    Public Sub FlyInFromLeft(X, y, w, h, alpha, pro As Single)
        X = X * pro
    End Sub
    Public Sub FlyInFromRight(X, y, w, h, alpha, pro As Single)
        X = X + (GW - X) * pro
    End Sub
    Public Sub FlyInFromTop(X, y, w, h, alpha, pro As Single)
        y = y * pro
    End Sub
    Public Sub FlyInFromBottom(X, y, w, h, alpha, pro As Single)
        y = y + (GH - y) * pro
    End Sub
    Public Sub FlyInFromLeftFade(X, y, w, h, alpha, pro As Single)
        X = X * pro
        alpha = alpha * pro
    End Sub
    Public Sub FlyInFromRightFade(X, y, w, h, alpha, pro As Single)
        X = X + (GW - X) * pro
        alpha = alpha * pro
    End Sub
    Public Sub FlyInFromTopFade(X, y, w, h, alpha, pro As Single)
        y = y * pro
        alpha = alpha * pro
    End Sub
    Public Sub FlyInFromBottomFade(X, y, w, h, alpha, pro As Single)
        y = y + (GH - y) * pro
        alpha = alpha * pro
    End Sub
'================================================================

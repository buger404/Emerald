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
    'Animation Project
    Public Sub FadeIn(X, Y, w, h, Alpha, pro As Single)
        Alpha = Alpha * pro
    End Sub
    Public Sub FadeOut(X, Y, w, h, Alpha, pro As Single)
        Alpha = Alpha - Alpha * pro
    End Sub
    Public Sub FlyInFromLeft(X, Y, w, h, Alpha, pro As Single)
        X = X * pro
    End Sub
    Public Sub FlyInFromRight(X, Y, w, h, Alpha, pro As Single)
        X = X + (GW - X) * pro
    End Sub
    Public Sub FlyInFromTop(X, Y, w, h, Alpha, pro As Single)
        Y = Y * pro
    End Sub
    Public Sub FlyInFromBottom(X, Y, w, h, Alpha, pro As Single)
        Y = Y + (GH - Y) * pro
    End Sub
    Public Sub FlyInFromLeftFade(X, Y, w, h, Alpha, pro As Single)
        X = X * pro
        Alpha = Alpha * pro
    End Sub
    Public Sub FlyInFromRightFade(X, Y, w, h, Alpha, pro As Single)
        X = X + (GW - X) * pro
        Alpha = Alpha * pro
    End Sub
    Public Sub FlyInFromTopFade(X, Y, w, h, Alpha, pro As Single)
        Y = Y * pro
        Alpha = Alpha * pro
    End Sub
    Public Sub FlyInFromBottomFade(X, Y, w, h, Alpha, pro As Single)
        Y = Y + (GH - Y) * pro
        Alpha = Alpha * pro
    End Sub
'================================================================

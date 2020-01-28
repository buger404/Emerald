VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GCrashBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim tF As GraphicsBound
Dim result As Boolean
Public Sub Reset()
    tF = DrawF
    result = False
End Sub
Public Function CrashResult() As Boolean
    CrashResult = result
End Function
Public Function IsCrash() As Boolean

    Dim StartX As Long, StartY As Long, EndX As Long, EndY As Long
    If tF.X > DrawF.X Then
        StartX = tF.X: EndX = DrawF.X + DrawF.Width
    Else
        StartX = DrawF.X: EndX = tF.X + tF.Width
    End If
    If tF.y > DrawF.y Then
        StartY = tF.y: EndY = DrawF.y + DrawF.Height
    Else
        StartY = DrawF.y: EndY = tF.y + tF.Height
    End If
    
    '碰都没碰到！GUN！
    If EndX - StartX < 0 Or EndY - StartY < 0 Then Exit Function
    
    '哦。那绝对撞到了。GUN！
    If DrawF.Shape <> -1 Then result = True: IsCrash = True: Exit Function
    
    If tF.CrashIndex = 0 Or DrawF.CrashIndex = 0 Then Exit Function
    
    Dim Ret As Boolean
    Dim X2 As Long, Y2 As Long
    Dim w1 As Long, w2 As Long, h1 As Long, h2 As Long
    w1 = UBound(ColorLists(tF.CrashIndex).IsAlpha, 1)
    h1 = UBound(ColorLists(tF.CrashIndex).IsAlpha, 2)
    w2 = UBound(ColorLists(DrawF.CrashIndex).IsAlpha, 1)
    h2 = UBound(ColorLists(DrawF.CrashIndex).IsAlpha, 2)
    
    For X = StartX - tF.X To EndX - tF.X
        For y = StartY - tF.y To EndY - tF.y
            If X * tF.WSc <= w1 And y * tF.HSc <= h1 And X >= 0 And y >= 0 Then
                If ColorLists(tF.CrashIndex).IsAlpha(X * tF.WSc, y * tF.HSc) Then
                    Ret = True
                    X2 = (X + tF.X - DrawF.X) * DrawF.WSc
                    Y2 = (y + tF.y - DrawF.y) * DrawF.HSc
                    If X2 <= w2 And Y2 <= h2 And X2 >= 0 And Y2 >= 0 Then
                        Ret = Ret And ColorLists(DrawF.CrashIndex).IsAlpha(X2, Y2)
                    Else
                        Ret = False
                    End If
                    '撞到我了！赔钱！没钱？GUN！
                    If Ret Then result = True: IsCrash = True: Exit Function
                End If
            End If
        Next
    Next
End Function

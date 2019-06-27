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
Dim Result As Boolean
Public Sub Reset()
    tF = DrawF
    Result = False
End Sub
Public Function CrashResult() As Boolean
    CrashResult = Result
End Function
Public Function IsCrash() As Boolean

    Dim StartX As Long, StartY As Long, EndX As Long, EndY As Long
    If tF.X > DrawF.X Then
        StartX = tF.X: EndX = DrawF.X + DrawF.Width
    Else
        StartX = DrawF.X: EndX = tF.X + tF.Width
    End If
    If tF.Y > DrawF.Y Then
        StartY = tF.Y: EndY = DrawF.Y + DrawF.Height
    Else
        StartY = DrawF.Y: EndY = tF.Y + tF.Height
    End If
    
    '碰都没碰到！GUN！
    If EndX - StartX < 0 Or EndY - StartY < 0 Then Exit Function
    
    '哦。那绝对撞到了。GUN！
    If DrawF.Shape <> -1 Then Result = True: IsCrash = True: Exit Function
    
    If tF.CrashIndex = 0 Or DrawF.CrashIndex = 0 Then Exit Function
    
    Dim ret As Boolean
    Dim x2 As Long, y2 As Long
    Dim w1 As Long, w2 As Long, h1 As Long, h2 As Long
    w1 = UBound(ColorLists(tF.CrashIndex).IsAlpha, 1)
    h1 = UBound(ColorLists(tF.CrashIndex).IsAlpha, 2)
    w2 = UBound(ColorLists(DrawF.CrashIndex).IsAlpha, 1)
    h2 = UBound(ColorLists(DrawF.CrashIndex).IsAlpha, 2)
    
    For X = StartX - tF.X To EndX - tF.X
        For Y = StartY - tF.Y To EndY - tF.Y
            If X <= w1 And Y <= h1 And X >= 0 And Y >= 0 Then
                If ColorLists(tF.CrashIndex).IsAlpha(X, Y) Then
                    ret = True
                    x2 = X + tF.X - DrawF.X
                    y2 = Y + tF.Y - DrawF.Y
                    If x2 <= w2 And y2 <= h2 And x2 >= 0 And y2 >= 0 Then
                        ret = ret And ColorLists(DrawF.CrashIndex).IsAlpha(x2, y2)
                    Else
                        ret = False
                    End If
                    '撞到我了！赔钱！没钱？GUN！
                    If ret Then Result = True: IsCrash = True: Exit Function
                End If
            End If
        Next
    Next
End Function

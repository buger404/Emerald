Attribute VB_Name = "Animations"
'==========================================================================
'   这里是Emerald游戏动画模块
'   嗯。如果我忘记写注释。希望未来的我不会在此表示批评。
'   未来的我：批评！
    Public Type EAniSound
        snd As String
        rate As Single
    End Type
    Public Type EAniFrame
        pic As String
        picindex As Integer
        x As Long
        y As Long
        size As Single
        alpha As Single
    End Type
    Public Type EAniTickFrame
        aframes() As EAniFrame
        sounds() As EAniSound
        tick As Long
        msg As String
        disposed As Boolean
    End Type
    Public Type EAniChannel
        name As String
        frames() As EAniTickFrame
        CurrentFrame As Integer
    End Type
    Public Type EAnimation
        channel() As EAniChannel
        CurrentChannel As Integer
        tick As Long
        globalTick As Long
        name As String
        position As PosAlign
    End Type
'==========================================================================
    Public Function LoadAnimation(path As String) As EAnimation
        Dim temp As String, ani As EAnimation, temp2() As String, temp3() As String
        Dim framepos As Long
        ReDim ani.channel(0) '初始化
        
        Open path For Input As #1
        Do While Not EOF(1)
            Line Input #1, temp
            temp = Trim(temp)
            '设置动画信息
            If ani.name = "" Then
                temp2 = Split(temp, "|")
                For i = 0 To UBound(temp2)
                    temp3 = Split(temp2(i), " ")
                    If temp3(0) = "name" Then ani.name = temp3(1)
                    If temp3(0) = "position" Then ani.position = Val(temp3(1))
                    If temp3(0) = "tick" Then ani.globalTick = Val(temp3(1))
                Next
                GoTo DoWithDone
            End If
            '发现通道头
            If Left(temp, 8) = ":Channel" Then
                ReDim Preserve ani.channel(UBound(ani.channel) + 1)
                With ani.channel(UBound(ani.channel))
                    ReDim .frames(0)
                    .name = Split(temp, ":Channel ")(1)
                End With
                GoTo DoWithDone
            End If
            '发现动画帧头
            If Left(temp, 1) = "{" Then
                With ani.channel(UBound(ani.channel))
                    ReDim Preserve .frames(UBound(.frames) + 1)
                    ReDim .frames(UBound(.frames)).aframes(0)
                    ReDim .frames(UBound(.frames)).sounds(0)
                    .frames(UBound(.frames)).tick = ani.globalTick
                    If Len(temp) > 1 Then
                        temp = Right(temp, Len(temp) - 1)
                        temp2 = Split(temp, "|")
                        For i = 0 To UBound(temp2)
                            temp3 = Split(temp2(i), " ")
                            If temp3(0) = "tick" Then .frames(UBound(.frames)).tick = Val(temp3(1))
                            If temp3(0) = "stay" Then .frames(UBound(.frames)).tick = 0
                            If temp3(0) = "dispose" Then .frames(UBound(.frames)).disposed = True
                            If temp3(0) = "msg" Then .frames(UBound(.frames)).msg = Split(temp2(i), """")(1)
                        Next
                    End If
                End With
                GoTo DoWithDone
            End If
            '发现动画帧尾
            If Left(temp, 1) = "}" Then
                framepos = 0
                GoTo DoWithDone
            End If
            '图片
            If framepos = 0 Then
                temp2 = Split(temp, "|")
                With ani.channel(UBound(ani.channel)).frames(UBound(ani.channel(UBound(ani.channel)).frames))
                    ReDim .aframes(UBound(temp2))
                    For i = 0 To UBound(temp2)
                        temp3 = Split(Right(Split(temp2(i), "(")(1), Len(temp2(i)) - 1), ",")
                        .aframes(i).size = Val(temp3(0))
                        .aframes(i).alpha = Val(temp3(1))
                        .aframes(i).x = Val(temp3(2))
                        .aframes(i).y = Val(temp3(3))
                        .aframes(i).pic = Split(temp2(i), "(")(0)
                    Next
                End With
            End If
            '音效
            If framepos = 1 Then
                temp2 = Split(temp, "|")
                With ani.channel(UBound(ani.channel)).frames(UBound(ani.channel(UBound(ani.channel)).frames))
                    ReDim .sounds(UBound(temp2))
                    For i = 0 To UBound(temp2)
                        temp3 = Split(temp2(i), "(")
                        temp3(1) = Right(temp3(1), Len(temp3(1)) - 1)
                        .sounds(i).snd = temp3(0)
                        .sounds(i).rate = Val(temp3(0))
                    Next
                End With
            End If
            framepos = framepos + 1
DoWithDone:
        Loop
        Close #1
        
        LoadAnimation = ani
    End Function
    

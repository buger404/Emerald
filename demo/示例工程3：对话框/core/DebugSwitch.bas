Attribute VB_Name = "DebugSwitch"
'   Emerald 设置项

'======================================================
'   相关设置请转到Builder中的“设置”
'======================================================
'   警告：不要修改下列代码
    Public DebugMode As Integer, DisableLOGO As Integer, HideLOGO As Integer, HideSuggest As Integer
    Public Debug_focus As Boolean, Debug_pos As Boolean, Debug_data As Boolean, Debug_mouse As Boolean, Debug_umode As Integer
    Public ChoosePosition As Boolean, ChooseRect As RECT
    Public ChooseLines() As ChooseLine
    Public FPSRecord(20) As Long
    Public Type ChooseLine
        mode As Integer
        Data As Long
        R As RECT
    End Type
    Public Sub JudgeChoosePosition(ByVal X As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
        If (Abs(Mouse.X - X) < 5) Then
            ReDim Preserve ChooseLines(UBound(ChooseLines) + 1)
            ChooseLines(UBound(ChooseLines)).mode = 0
            ChooseLines(UBound(ChooseLines)).Data = X
            With ChooseLines(UBound(ChooseLines)).R
                .Left = X: .top = y: .Right = w: .Bottom = h
            End With
        End If
        If (Abs(Mouse.y - y) < 5) Then
            ReDim Preserve ChooseLines(UBound(ChooseLines) + 1)
            ChooseLines(UBound(ChooseLines)).mode = 1
            ChooseLines(UBound(ChooseLines)).Data = y
            With ChooseLines(UBound(ChooseLines)).R
                .Left = X: .top = y: .Right = w: .Bottom = h
            End With
        End If
        If (Abs(Mouse.X - (X + w)) < 5) Then
            ReDim Preserve ChooseLines(UBound(ChooseLines) + 1)
            ChooseLines(UBound(ChooseLines)).mode = 0
            ChooseLines(UBound(ChooseLines)).Data = X + w
            With ChooseLines(UBound(ChooseLines)).R
                .Left = X: .top = y: .Right = w: .Bottom = h
            End With
        End If
        If (Abs(Mouse.y - (y + h)) < 5) Then
            ReDim Preserve ChooseLines(UBound(ChooseLines) + 1)
            ChooseLines(UBound(ChooseLines)).mode = 1
            ChooseLines(UBound(ChooseLines)).Data = y + h
            With ChooseLines(UBound(ChooseLines)).R
                .Left = X: .top = y: .Right = w: .Bottom = h
            End With
        End If
    End Sub
'======================================================

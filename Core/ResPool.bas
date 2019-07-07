Attribute VB_Name = "ResPool"
Dim DC() As Long, brush() As Long, Font() As Long, StrF() As Long, Pen() As Long
Dim Graphics() As Long, Image() As Long, Effect() As Long, Path() As Long
Dim Obj() As Object
Public Function GetCountStr() As String
    GetCountStr = "DC " & UBound(DC) & " , Brush " & UBound(brush) & " , Pen " & UBound(Pen) & vbCrLf & _
                  "Font " & UBound(Font) & " , StringFormat " & UBound(StrF) & vbCrLf & _
                  "Bitmap " & UBound(Image) & " , Object " & UBound(Obj) & vbCrLf & _
                  "Effect " & UBound(Effect) & " , Graphics " & UBound(Graphics) & " , Path " & UBound(Path)
End Function
Public Sub InitPool()
    ReDim DC(0)
    ReDim brush(0)
    ReDim Font(0)
    ReDim StrF(0)
    ReDim Pen(0)
    ReDim Graphics(0)
    ReDim Image(0)
    ReDim Obj(0)
    ReDim Effect(0)
    ReDim Path(0)
    If App.LogMode = 0 Then Open VBA.Environ("temp") & "\Emerald " & year(Now) & "_" & Month(Now) & "_" & Day(Now) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & "_" & App.ThreadID & ".txt" For Output As #446
End Sub
Public Sub EmrLog(Str As String)
    If App.LogMode = 0 Then Print #446, Now & "   " & Str
End Sub
Public Sub DestroyPool()
    On Error Resume Next
    EmrLog "Emerald ResPool Version " & Version
    EmrLog "Work starting ..."
    For i = 1 To UBound(DC)
        If DC(i) <> 0 Then DeleteObject DC(i): EmrLog "Delete DC " & DC(i)
    Next
    For i = 1 To UBound(brush)
        If brush(i) <> 0 Then gdiplus.GdipDeleteBrush brush(i): EmrLog "Delete Brush " & brush(i)
    Next
    For i = 1 To UBound(Pen)
        If Pen(i) <> 0 Then gdiplus.GdipDeletePen Pen(i): EmrLog "Delete Pen " & Pen(i)
    Next
    For i = 1 To UBound(Graphics)
        If Graphics(i) <> 0 Then gdiplus.GdipDeleteGraphics Graphics(i): EmrLog "Delete Graphics " & Graphics(i)
    Next
    For i = 1 To UBound(Image)
        If Image(i) <> 0 Then gdiplus.GdipDisposeImage Image(i): EmrLog "Delete Image " & Image(i)
    Next
    For i = 1 To UBound(Font)
        If Font(i) <> 0 Then gdiplus.GdipDeleteFont Font(i): EmrLog "Delete Font " & Font(i)
    Next
    For i = 1 To UBound(StrF)
        If StrF(i) <> 0 Then gdiplus.GdipDeleteStringFormat StrF(i): EmrLog "Delete StrF " & StrF(i)
    Next
    For i = 1 To UBound(Effect)
        If Effect(i) <> 0 Then gdiplus.GdipDeleteEffect Effect(i): EmrLog "Delete Effect " & Effect(i)
    Next
    For i = 1 To UBound(Path)
        If Path(i) <> 0 Then gdiplus.GdipDeletePath Path(i): EmrLog "Delete Path " & Path(i)
    Next
    For i = 1 To UBound(Obj)
        If Not Obj(i) Is Nothing Then EmrLog "Delete Object " & ObjPtr(Obj(i)): Set Obj(i) = Nothing
    Next
    
    If App.LogMode = 0 Then Close #446
End Sub
Public Sub DeleteObj(Hwnd As Long)
    For i = 1 To UBound(Obj)
        If Obj(i) = Hwnd Then Set Obj(i) = Nothing: Exit For
    Next
End Sub
Public Sub PoolDisposeImage(Hwnd As Long)
    For i = 1 To UBound(Image)
        If Image(i) = Hwnd Then Image(i) = 0: Exit For
    Next
    gdiplus.GdipDisposeImage Hwnd
End Sub
Public Sub PoolDeletePath(Hwnd As Long)
    For i = 1 To UBound(Path)
        If Path(i) = Hwnd Then Path(i) = 0: Exit For
    Next
    gdiplus.GdipDeletePath Hwnd
End Sub
Public Sub PoolDeleteEffect(Hwnd As Long)
    For i = 1 To UBound(Effect)
        If Effect(i) = Hwnd Then Effect(i) = 0: Exit For
    Next
    gdiplus.GdipDeleteEffect Hwnd
End Sub
Public Sub PoolDeleteGraphics(Hwnd As Long)
    For i = 1 To UBound(Graphics)
        If Graphics(i) = Hwnd Then Graphics(i) = 0: Exit For
    Next
    gdiplus.GdipDeleteGraphics Hwnd
End Sub
Public Sub PoolDeletePen(Hwnd As Long)
    For i = 1 To UBound(Pen)
        If Pen(i) = Hwnd Then Pen(i) = 0: Exit For
    Next
    gdiplus.GdipDeletePen Hwnd
End Sub
Public Sub PoolDeleteStringFormat(Hwnd As Long)
    For i = 1 To UBound(StrF)
        If StrF(i) = Hwnd Then StrF(i) = 0: Exit For
    Next
    gdiplus.GdipDeleteStringFormat Hwnd
End Sub
Public Sub PoolDeleteFont(Hwnd As Long)
    For i = 1 To UBound(Font)
        If Font(i) = Hwnd Then Font(i) = 0: Exit For
    Next
    gdiplus.GdipDeleteFont Hwnd
End Sub
Public Sub PoolDeleteBrush(Hwnd As Long)
    For i = 1 To UBound(brush)
        If brush(i) = Hwnd Then brush(i) = 0: Exit For
    Next
    gdiplus.GdipDeleteBrush Hwnd
End Sub
Public Sub DeleteDC(Hwnd As Long)
    For i = 1 To UBound(DC)
        If DC(i) = Hwnd Then DC(i) = 0: Exit For
    Next
    DeleteObject Hwnd
End Sub
Public Sub PoolAddObject(nObj As Object)
    ReDim Preserve Obj(UBound(Obj) + 1)
    Set Obj(UBound(Obj)) = nObj
End Sub
Public Sub PoolAddPath(Hwnd As Long)
    ReDim Preserve Path(UBound(Path) + 1)
    Path(UBound(Path)) = Hwnd
End Sub
Public Sub PoolAddEffect(Hwnd As Long)
    ReDim Preserve Effect(UBound(Effect) + 1)
    Effect(UBound(Effect)) = Hwnd
End Sub
Public Sub PoolAddDC(Hwnd As Long)
    ReDim Preserve DC(UBound(DC) + 1)
    DC(UBound(DC)) = Hwnd
End Sub
Public Sub PoolAddBrush(Hwnd As Long)
    ReDim Preserve brush(UBound(brush) + 1)
    brush(UBound(brush)) = Hwnd
End Sub
Public Sub PoolAddFont(Hwnd As Long)
    ReDim Preserve Font(UBound(Font) + 1)
    Font(UBound(Font)) = Hwnd
End Sub
Public Sub PoolAddStrF(Hwnd As Long)
    ReDim Preserve StrF(UBound(StrF) + 1)
    StrF(UBound(StrF)) = Hwnd
End Sub
Public Sub PoolAddPen(Hwnd As Long)
    ReDim Preserve Pen(UBound(Pen) + 1)
    Pen(UBound(Pen)) = Hwnd
End Sub
Public Sub PoolAddGraphics(Hwnd As Long)
    ReDim Preserve Graphics(UBound(Graphics) + 1)
    Graphics(UBound(Graphics)) = Hwnd
End Sub
Public Sub PoolAddImage(Hwnd As Long)
    ReDim Preserve Image(UBound(Image) + 1)
    Image(UBound(Image)) = Hwnd
End Sub
Public Sub PoolCreateStringFormat(atr As Long, lan As Integer, format As Long)
    gdiplus.GdipCreateStringFormat atr, lan, format
    PoolAddStrF format
End Sub
Public Sub PoolCreateFont(fam As Long, ByVal size As Single, style As FontStyle, unit As GpUnit, Font As Long)
    gdiplus.GdipCreateFont fam, size, style, unit, Font
    PoolAddFont Font
End Sub
Public Sub PoolCreateEffect2(eff As GdipEffectType, Effect As Long)
    gdiplus.GdipCreateEffect2 eff, Effect
    PoolAddEffect Effect
End Sub
Public Function PoolCreateObject(Str As String) As Object
    Set PoolCreateObject = VBA.CreateObject(Str)
    PoolAddObject PoolCreateObject
End Function
Public Sub PoolCreateBitmapFromFile(filename As Long, bitmap As Long)
    gdiplus.GdipCreateBitmapFromFile filename, bitmap
    PoolAddImage bitmap
End Sub
Public Sub PoolCreatePen1(argb As Long, Width As Single, unit As GpUnit, Pen As Long)
    gdiplus.GdipCreatePen1 argb, Width, unit, Pen
    PoolAddPen Pen
End Sub
Public Sub PoolCreateSolidFill(argb As Long, brush As Long)
    gdiplus.GdipCreateSolidFill argb, brush
    PoolAddBrush brush
End Sub
Public Sub PoolCreatePath(mode As FillMode, Path As Long)
    gdiplus.GdipCreatePath mode, Path
    PoolAddPath Path
End Sub
Public Sub PoolCreateFromHdc(DC As Long, g As Long)
    gdiplus.GdipCreateFromHDC DC, g
    PoolAddGraphics g
End Sub
Public Function CreateCDC(w As Long, h As Long) As Long
    Dim bm As BITMAPINFOHEADER, DC As Long, DIB As Long

    With bm
        .biBitCount = 32
        .biHeight = h
        .biWidth = w
        .biPlanes = 1
        .biSizeImage = (.biWidth * .biBitCount + 31) / 32 * 4 * .biHeight
        .biSize = Len(bm)
    End With
    
    DC = CreateCompatibleDC(GDC)
    DIB = CreateDIBSection(DC, bm, DIB_RGB_COLORS, ByVal 0, 0, 0)
    DeleteObject SelectObject(DC, DIB)
    DeleteObject DIB
    
    CreateCDC = DC
    PoolAddDC DC
End Function

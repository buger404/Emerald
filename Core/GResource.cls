VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As clsid) As Long

Dim Mems() As GMem

Private Sub Class_Initialize()
    ReDim Mems(0)
End Sub
Public Sub Add(mem As Long, kind As Integer, name As String, Optional w, Optional h, Optional folder As String = "")
    ReDim Preserve Mems(UBound(Mems) + 1)
    With Mems(UBound(Mems))
        .Hwnd = mem
        .kind = kind
        .name = name
        .folder = folder
        If Not IsMissing(w) Then .w = w
        If Not IsMissing(h) Then .h = h
    End With
End Sub
Public Property Get ImgCount() As Long
    ImgCount = UBound(Mems)
End Property
Public Property Get ImgSize(n, sizeIndex As imgIndex)
    Dim index As Integer
    
    If TypeName(n) = "String" Then
        index = GetImage(n)
    Else
        index = n
    End If

    Select Case sizeIndex
        Case 0
            ImgSize = Mems(index).w
        Case 1
            ImgSize = Mems(index).h
        Case 2
            ImgSize = Mems(index).GIF.Count
    End Select
End Property
Public Sub Dispose()
    For i = 1 To UBound(Mems)
        If Not Mems(i).copyed Then
            Select Case Mems(i).kind
                Case 0 'DC
                    DeleteObject Mems(i).Hwnd
                    GdipDisposeImage Mems(i).ImgHwnd
                Case 1 'Graphics
                    GdipDeleteGraphics Mems(i).Hwnd
            End Select
            If Mems(i).GIF.Count > 0 Then
                For s = 0 To Mems(i).GIF.Count
                    DeleteObject Mems(i).GIF.frames(s)
                Next
            End If
        End If
    Next
End Sub
Public Property Get ResourceGIFHandle(index As Integer, frame As Integer) As Long
    ResourceGIFHandle = Mems(index).GIF.frames(frame)
End Property
Public Property Get ResourceGIFCurrentHandle(index As Integer) As Long
    ResourceGIFCurrentHandle = Mems(index).GIF.frames(Mems(index).GIF.tick)
End Property
Public Property Get ResourceGIFTick(index As Integer) As Long
    ResourceGIFTick = Mems(index).GIF.tick
End Property
Public Sub GIFTick(index As Integer)
    Mems(index).GIF.tick = Mems(index).GIF.tick + 1
    If Mems(index).GIF.tick > Mems(index).GIF.Count Then Mems(index).GIF.tick = 0
End Sub
Public Property Get ResourceGIFCount(index As Integer) As Long
    ResourceGIFCount = Mems(index).GIF.Count
End Property
Public Property Get ResourceGIFTime(index As Integer) As Long
    ResourceGIFTime = Mems(index).GIF.time
End Property
Public Property Get ResourceName(index As Integer) As Long
    ResourceName = Mems(index).name
End Property
Public Property Get ResourceHandle(index As Integer) As Long
    ResourceHandle = Mems(index).Hwnd
End Property
Public Property Get ResourceImageHandle(index As Integer) As Long
    ResourceImageHandle = Mems(index).ImgHwnd
End Property
Public Function GetImage(ByVal name As String) As Integer
    For i = 1 To UBound(Mems)
        If Mems(i).kind = 0 Then
            If Mems(i).name = name Then GetImage = i: Exit For
        End If
    Next
End Function
Public Sub newImage(Path As String, Optional arg1, Optional arg2)
    'when arg1 is missing ; i.w = i.w
    'when arg2 is missing but arg1 has content ; i.w = i.w * arg1
    'other ; i.w = arg1
    Dim si As Integer, i As Long, w As Long, h As Long, DC As Long, g As Long
    If IsMissing(arg1) Then si = 0
    If IsMissing(arg2) And (Not IsMissing(arg1)) Then si = 1
    If (Not IsMissing(arg2)) And (Not IsMissing(arg1)) Then si = 2
    
    GdipCreateBitmapFromFile StrPtr(Path), i
    GdipGetImageWidth i, w: GdipGetImageHeight i, h
    
    If si = 1 Then w = Round(w * arg1): h = Round(h * arg1)
    If si = 2 Then w = arg1: h = arg2
    
    Dim gC As clsid, gL As Long, index As Long, temp() As String
    CLSIDFromString StrPtr("{6AEDBD6D-3FB5-418A-83A6-7F45229DC872}"), gC
    GdipImageGetFrameCount i, gC, gL
    If gL > 0 Then  '¶¯Ì¬GIFÖ§³Ö
        temp = Split(Path, "\")
        Add 0, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
        With Mems(UBound(Mems))
            .GIF.Count = gL
            ReDim .GIF.frames(gL)
            For index = 0 To gL
                GdipImageSelectActiveFrame i, gC, index
                DC = CreateCDC(w, h)
                GdipCreateFromHDC DC, g
                GdipDrawImageRect g, i, 0, 0, w, h
                GdipDeleteGraphics g
                .GIF.frames(index) = DC
            Next
        End With
        
        Mems(UBound(Mems)).ImgHwnd = i
        Exit Sub
    End If
    
    DC = CreateCDC(w, h)
    GdipCreateFromHDC DC, g
    GdipDrawImageRect g, i, 0, 0, w, h
    
    GdipDeleteGraphics g
    
    temp = Split(Path, "\")
    Add DC, 0, temp(UBound(temp)), w, h, temp(UBound(temp) - 1)
    Mems(UBound(Mems)).ImgHwnd = i
End Sub
Public Sub NewImages(Path As String, Optional arg1, Optional arg2)
    Dim f As String, c As Long
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    Dim index As Integer, Tree As AssetsTree
    Tree.Path = Path: ReDim Tree.Files(0)
    index = FindAssetsTree(Path)
    
    f = Dir(Path)
    Do While f <> ""
        c = c + 1
        ReDim Preserve Tree.Files(UBound(Tree.Files) + 1)
        f = Dir()
    Loop
    If index = 0 Then
        PreLoadCount = PreLoadCount + c
    Else
        Tree = GetAssetsTree(Path)
        ReLoadCount = ReLoadCount + UBound(Tree.Files)
        For i = 1 To UBound(Tree.Files)
            ReDim Preserve Mems(UBound(Mems) + 1)
            Mems(UBound(Mems)) = Tree.Files(i)
            Mems(UBound(Mems)).copyed = True
        Next
        Exit Sub
    End If
    
    f = Dir(Path)
    i = 1
    Do While f <> ""
        If IsMissing(arg1) Then
            newImage Path & f
        ElseIf IsMissing(arg2) Then
            newImage Path & f, arg1
        Else
            newImage Path & f, arg1, arg2
        End If
        Tree.Files(i) = Mems(UBound(Mems))
        i = i + 1
        f = Dir()
        DoEvents
        LoadedCount = LoadedCount + 1
    Loop
    
    If index = 0 Then AddAssetsTree Tree
End Sub


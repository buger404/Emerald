VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Emerald Ïà¹Ø´úÂë

Dim FFTData(512) As Single, FFTPeacks(128) As Long, FFTFallOff(128) As Long
Dim Hwnd As Long, mRate As Long, mVolume As Single
'========================================================
'   Properties
    Public Property Get PlayState() As PlayStateMark
        PlayState = BASS_ChannelIsActive(Hwnd)
    End Property
    Public Property Get IsPlaying() As Boolean
        IsPlaying = (BASS_ChannelIsActive(Hwnd) <> 0)
    End Property
    Public Property Get length() As Double
        length = BASS_ChannelBytes2Seconds(Hwnd, BASS_ChannelGetLength(Hwnd, BASS_POS_BYTE))
    End Property
    Public Property Get position() As Double
        position = BASS_ChannelBytes2Seconds(Hwnd, BASS_ChannelGetPosition(Hwnd, BASS_POS_BYTE))
    End Property
    Public Property Let position(p As Double)
        BASS_ChannelSetPosition Hwnd, BASS_ChannelSeconds2Bytes(Hwnd, p), BASS_POS_BYTE
    End Property
    Public Property Get Volume() As Single
        Volume = mVolume
    End Property
    Public Property Let Volume(v As Single)
        BASS_ChannelSetAttribute Hwnd, BASS_ATTRIB_VOL, v
        mVolume = v
    End Property
    Public Property Get rate() As Single
        rate = mRate
    End Property
    Public Property Let rate(r As Single)
        BASS_ChannelSetAttribute Hwnd, BASS_ATTRIB_FREQ, 44100 * r
        mRate = r
    End Property
'========================================================
'   Init
    Private Sub Class_Initialize()
        mVolume = 0.5: mRate = 1
    End Sub
    Public Sub Create(ByVal path As String)
        If Not BassInstalled Then
            BASS_Init -1, 44100, BASS_DEVICE_3D, Hwnd, 0
            BassInstalled = True
        End If
        
        If Hwnd <> 0 Then Call Dispose
        Hwnd = BASS_StreamCreateFile(BASSFALSE, StrPtr(path), 0, 0, 0)
        BASS_ChannelSetAttribute Hwnd, BASS_ATTRIB_VOL, mVolume
        BASS_ChannelSetAttribute Hwnd, BASS_ATTRIB_FREQ, 44100 * mRate
    End Sub
    Public Sub Dispose()
        Call Stops
        BASS_StreamFree Hwnd
    End Sub
'========================================================
'   Control
    Public Sub Continue()
        Call BASS_ChannelPlay(Hwnd, BASSFALSE)
    End Sub
    Public Sub Play()
        Call BASS_ChannelPlay(Hwnd, BASSTRUE)
    End Sub
    Public Sub Stops()
        Call BASS_ChannelStop(Hwnd)
    End Sub
    Public Sub Pause()
        Call BASS_ChannelPause(Hwnd)
    End Sub
'========================================================
'   MusicMap
    Function GetMusicMap() As Single()
        On Error Resume Next
        Dim i As Long, di As Long, j As Long, returns(512) As Single
        Dim tmp As Long
        
        If BASS_ChannelIsActive(Hwnd) <> BASS_ACTIVE_PLAYING Then Exit Function
        
        Call BASS_ChannelGetData(Hwnd, FFTData(0), BASS_DATA_FFT1024)
        
        For i = 0 To UBound(FFTData) - 1
            di = Int(Abs(FFTData(i)) * 7500)
            j = i Mod 129
            If di > 1500 Then di = 1500
            If di >= FFTPeacks(j) Then
                FFTPeacks(j) = di
            Else
                FFTPeacks(j) = FFTPeacks(j) - 10
            End If
            If di >= FFTFallOff(j) Then
                FFTFallOff(j) = di
            Else
                FFTFallOff(j) = FFTFallOff(j) - 30
            End If
            If (1500 - FFTPeacks(j)) > 1500 Then FFTPeacks(j) = 0
            If (1500 - FFTFallOff(j)) > 1500 Then FFTFallOff(j) = 0
            returns(i) = FFTFallOff(j)
        Next
        GetMusicMap = returns
    End Function
'========================================================

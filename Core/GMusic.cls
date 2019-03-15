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
Public Enum PlayState
    Stopped = 0
    Playing = 1
    Stalled = 2
    Paused = 3
End Enum
Dim hwnd As Long, mRate As Long, mVolume As Single
'========================================================
'   Properties
    Public Property Get Length() As Double
        Length = BASS_ChannelBytes2Seconds(hwnd, BASS_ChannelGetLength(SongHandle, BASS_POS_BYTE))
    End Property
    Public Property Get Position() As Double
        Position = BASS_ChannelBytes2Seconds(hwnd, BASS_ChannelGetPosition(SongHandle, BASS_POS_BYTE))
    End Property
    Public Property Let Position(p As Double)
        BASS_ChannelSetPosition hwnd, BASS_ChannelSeconds2Bytes(hwnd, p), BASS_POS_BYTE
    End Property
    Public Property Get Volume() As Single
        Volume = mVolume
    End Property
    Public Property Let Volume(v As Single)
        BASS_ChannelSetAttribute hwnd, BASS_ATTRIB_VOL, v
        mVolume = v
    End Property
    Public Property Get Rate() As Single
        Rate = mRate
    End Property
    Public Property Let Rate(r As Single)
        BASS_ChannelSetAttribute hwnd, BASS_ATTRIB_FREQ, 44100 * r
        mRate = r
    End Property
'========================================================
'   Init
    Private Sub Class_Initialize()
        mVolume = 0.5: mRate = 1
    End Sub
    Public Sub Create(ByVal path As String)
        If hwnd <> 0 Then Call Dispose
        hwnd = BASS_StreamCreateFile(BASSFALSE, StrPtr(path), 0, 0, 0)
        BASS_ChannelSetAttribute hwnd, BASS_ATTRIB_VOL, mVolume
        BASS_ChannelSetAttribute hwnd, BASS_ATTRIB_FREQ, 44100 * mRate
    End Sub
    Public Sub Dispose()
        Call Stops
        BASS_StreamFree hwnd
    End Sub
'========================================================
'   Control
    Public Sub Continue()
        Call BASS_ChannelPlay(hwnd, BASSFALSE)
    End Sub
    Public Sub Play()
        Call BASS_ChannelPlay(hwnd, BASSTRUE)
    End Sub
    Public Sub Stops()
        Call BASS_ChannelStop(hwnd)
    End Sub
    Public Sub Pause()
        Call BASS_ChannelPause(hwnd)
    End Sub
'========================================================

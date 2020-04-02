Attribute VB_Name = "AeroEffect"
'Emerald Ïà¹Ø´úÂë

Private Declare Function DwmIsCompositionEnabled Lib "DwmApi.dll" (ByRef Enabled As Boolean) As Long
Private Declare Function DwmEnableComposition Lib "DwmApi.dll" (ByVal compositionAction As CompositionEnable) As Long

Private Declare Function DwmExtendFrameIntoClientArea Lib "DwmApi.dll" (ByVal Hwnd As Long, ByRef m As MARGINS) As Long
Private Declare Function DwmEnableBlurBehindWindow Lib "DwmApi.dll" (ByVal Hwnd As Long, ByRef bb As DWM_BLURBEHIND) As Long

Private Enum CompositionEnable
    DWM_EC_DISABLECOMPOSITION = 0
    DWM_EC_ENABLECOMPOSITION = 1
End Enum

Private Enum DwmBlurBehindDwFlags
    DWM_BB_ENABLE = 1
    DWM_BB_BLURREGION = 2
    DWM_BB_TRANSITIONONMAXIMIZED = 4
End Enum

Private Type DWM_BLURBEHIND
    dwFlags As DwmBlurBehindDwFlags
    fEnable As Boolean
    hRgnBlur As Long
    fTransitionOnMaximized As Boolean
End Type

Private Type MARGINS
    cxLeftWidth As Long
    cxRightWidth As Long
    cyBottomHeight As Long
    cyTopHeight As Long
End Type

Public Declare Function SetWindowCompositionAttribute Lib "user32.dll" (ByVal Hwnd As Long, ByRef Data As WindowsCompostionAttributeData) As Long

Public Enum WindowCompositionAttribute
    WCA_ACCENT_POLICY = 19
End Enum

Public Type WindowsCompostionAttributeData
    Attribute As WindowCompositionAttribute
    Data As Long
    SizeOfData As Long
End Type

Public Enum AccentState
    ACCENT_DISABLED = 0
    ACCENT_ENABLE_GRADIENT = 1
    ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
    ACCENT_ENABLE_BLURBEHIND = 3
    ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    ACCENT_INVALID_STATE = 5
End Enum

Public Type AccentPolicy
    State As AccentState
    flags As Long
    GradientColor As Long
    id As Long
End Type
Public Sub Win10Blur(Hwnd As Long, Color As Long)
    Dim Accent As AccentPolicy, Data As WindowsCompostionAttributeData
    
    With Accent
        .State = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND
        .GradientColor = Color
    End With
    
    With Data
        .Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY
        .SizeOfData = 16
        .Data = VarPtr(Accent)
    End With
    
    SetWindowCompositionAttribute ByVal Hwnd, Data
    
End Sub

Sub Win7Aeros(Hwnd As Long)
    Dim B As DWM_BLURBEHIND
    B.dwFlags = DWM_BB_ENABLE
    B.fEnable = True
    B.fTransitionOnMaximized = True
    B.hRgnBlur = vbNull
    DwmEnableBlurBehindWindow Hwnd, B
End Sub

Sub BlurWindow(Hwnd As Long, Optional BlurColor As Long = -1)
    Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    If BlurColor = -1 Then BlurColor = argb(120, 64, 64, 72)
    
    For Each objItem In colItems
        strOSversion = objItem.Version
    Next
    
    Select Case Left(strOSversion, 3)
    Case "10."                                              'Windows 10
        osver = Split(strOSversion, ".")
        If Val(osver(2)) >= 15063 Then Win10Blur Hwnd, BlurColor
    Case "6.1"                                              'Windows 7
        Win7Aeros Hwnd
    Case Else                                               'Dont Blur
        Exit Sub
    End Select
End Sub

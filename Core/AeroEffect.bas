Attribute VB_Name = "AeroEffect"
'Emerald ��ش���

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

Enum WindowCompositionAttribute
    WCA_ACCENT_POLICY = 19
End Enum

Type WindowsCompostionAttributeData
    Attribute As WindowCompositionAttribute
    Data As Long
    SizeOfData As Integer
End Type

Enum AccentState
    ACCENT_DISABLED = 0
    ACCENT_ENABLE_GRADIENT = 1
    ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
    ACCENT_ENABLE_BLURBEHIND = 3
    ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
    ACCENT_INVALID_STATE = 5
End Enum

Type AccentPolicy
    State As AccentState
    flags As Integer
    GradientColor As Integer
    id As Integer
End Type

Public Sub Win10Blur(Hwnd As Long)
    Dim Accent As AccentPolicy
    Accent.State = 3
    
    Dim AccentStructSize As Long
    AccentStructSize = 16
    
    Dim AccentPtr As Long
    AccentPtr = VarPtr(Accent)
    
    Dim Data As WindowsCompostionAttributeData
    With Data
        .Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY
        .SizeOfData = 16
        .Data = AccentPtr
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

Sub BlurWindow(Hwnd As Long)
    Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each objItem In colItems
        strOSversion = objItem.Version
    Next
    
    Select Case Left(strOSversion, 3)
    Case "10."                                              'Windows 10
        osver = Split(strOSversion, ".")
        If Val(osver(2)) >= 15063 Then Win10Blur Hwnd
    Case "6.1"                                              'Windows 7
        Win7Aeros Hwnd
    Case Else                                               'Dont Blur
        Exit Sub
    End Select
End Sub

Attribute VB_Name = "SpecialDirs"
Public Declare Function SHGetSpecialFolderLocation Lib "Shell32" (ByVal hwndOwner As Long, ByVal nFolder As Integer, ppidl As Long) As Long
Public Declare Function SHGetPathFromIDList Lib "Shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal szPath As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Enum SpecialDir
    DESKTOP = &H0& '桌面
    PROGRAMS = &H2& '程序集
    MYDOCUMENTS = &H5& '我的文档
    MYFAVORITES = &H6& '收藏夹
    STARTUP = &H7& '启动
    RECENT = &H8& '最近打开的文件
    SENDTO = &H9& '发送
    STARTMENU = &HB& '开始菜单
    NETHOOD = &H13& '网上邻居
    FONTS = &H14& '字体
    SHELLNEW = &H15& 'ShellNew
    APPDATA = &H1A& 'ApplicationData
    PRINTHOOD = &H1B& 'PrintHood
    PAGETMP = &H20& '网页临时文件
    COOKIES = &H21& 'Cookies目录
    HISTORY = &H22& '历史
End Enum
Function GetSpecialDir(Dirs As SpecialDir) As String
    Dim sTmp As String * 200, nLength As Long, pidl As Long
    SHGetSpecialFolderLocation 0, Dirs, pidl
    SHGetPathFromIDList pidl, sTmp
    GetSpecialDir = Left(sTmp, InStr(sTmp, Chr(0)) - 1)
End Function

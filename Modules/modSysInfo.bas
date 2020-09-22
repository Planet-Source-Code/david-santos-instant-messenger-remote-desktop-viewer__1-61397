Attribute VB_Name = "modSysInfo"
Option Explicit

Public Enum CONST_CSIDL
    '95/98 Systems
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_ALTSTARTUP = &H1D                          'DBCS
    CSIDL_COMMON_ALTSTARTUP = &H1E                   'DBCS
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    
    'NT/XP SYSTEMS
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEOS = &HE
    CSIDL_LOCALAPPDATA = &H1C
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMDIRECTORY = &H29
    CSIDL_COMMON_FILES = &H2B
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_MYMUSIC = &H35
    CSIDL_COMMON_MYPICTURES = &H36
    CSIDL_COMMON_MYVIDEOS = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_CDBURNING = &H3B
End Enum


Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function GetSystemFolder(foldertype As CONST_CSIDL) As String
Dim strFolder As String
Dim lngIDL As Long

    strFolder = String(255, vbNullChar)

    If SHGetSpecialFolderLocation(frmMain.hwnd, foldertype, lngIDL) = 0 Then
        
       If SHGetPathFromIDList(lngIDL, strFolder) Then
           strFolder = Left(strFolder, InStr(1, strFolder, vbNullChar) - 1)
       Else
           strFolder = ""
       End If
    Else
        strFolder = ""
    
    End If
    
    GetSystemFolder = strFolder
End Function



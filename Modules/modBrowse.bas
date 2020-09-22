Attribute VB_Name = "modBrowse"
Option Explicit

Private Type BROWSEINFO
   hwndOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   pszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Const MAX_PATH As Long = 260
Private Const dhcErrorExtendedError = 1208&
Private Const dhcNoError = 0&

Public Enum CONSTSTARTBROWSE
 dhcCSIdlDesktop = &H0
 dhcCSIdlPrograms = &H2
 dhcCSIdlControlPanel = &H3
 dhcCSIdlInstalledPrinters = &H4
 dhcCSIdlPersonal = &H5
 dhcCSIdlFavorites = &H6
 dhcCSIdlStartupPmGroup = &H7
 dhcCSIdlRecentDocDir = &H8
 dhcCSIdlSendToItemsDir = &H9
 dhcCSIdlRecycleBin = &HA
 dhcCSIdlStartMenu = &HB
 dhcCSIdlDesktopDirectory = &H10
 dhcCSIdlMyComputer = &H11
 dhcCSIdlNetworkNeighborhood = &H12
 dhcCSIdlNetHoodFileSystemDir = &H13
 dhcCSIdlFonts = &H14
 dhcCSIdlTemplates = &H15
End Enum

'private constants for limiting choices for BrowseForFolder Dialog

Public Enum CONSTBROWSEBEHAVIOUR
 dhcBifReturnAll = &H0
 dhcBifReturnOnlyFileSystemDirs = &H1
 dhcBifDontGoBelowDomain = &H2
 dhcBifIncludeStatusText = &H4
 dhcBifSystemAncestors = &H8
 dhcBifBrowseForComputer = &H1000
 dhcBifBrowseForPrinter = &H2000
End Enum

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'Default Property Values:
Const m_def_Text = "Choose a folder"
Const m_def_LimitTo = 0
Const m_def_StartAt = 0
Const m_def_CancelError = False
'Property Variables:
Dim m_Text As String
Dim m_LimitTo As Long
Dim m_StartAt As Variant
Dim m_CancelError As Boolean

Private Function BrowseForFolder(ByVal lngCSIDL As CONSTSTARTBROWSE, _
   ByVal lngBiFlags As CONSTBROWSEBEHAVIOUR, _
   strFolder As String, _
   Optional ByVal hWnd As Long = 0, _
   Optional pszTitle As String = "Select Folder") As Long

Dim usrBrws As BROWSEINFO
Dim lngReturn As Long
Dim lngIDL As Long

If SHGetSpecialFolderLocation(hWnd, lngCSIDL, lngIDL) = 0 Then

   'set up the browse structure here
   With usrBrws
       .hwndOwner = hWnd
       .pidlRoot = lngIDL
       .pszDisplayName = String$(MAX_PATH, vbNullChar)
       .pszTitle = pszTitle
       .ulFlags = lngBiFlags
   End With

   'open the dialog
   lngIDL = SHBrowseForFolder(usrBrws)

   'if successful
       If lngIDL = 0 Then
           strFolder = ""
       Else
           strFolder = String$(MAX_PATH, vbNullChar)
       End If
       
       'resolve the long value form the lngIDL to a real path
       If SHGetPathFromIDList(lngIDL, strFolder) Then
           strFolder = Left(strFolder, InStr(1, strFolder, vbNullChar))
       lngReturn = dhcNoError 'to show there is no error.
       Else
           'nothing real is available.
           'return a virtual selection
           strFolder = Left(usrBrws.pszDisplayName, InStr(1, usrBrws.pszDisplayName, vbNullChar))
       lngReturn = dhcNoError 'to show there is no error.
       End If
Else
   lngReturn = dhcErrorExtendedError 'something went wrong
End If


BrowseForFolder = lngReturn

End Function

Public Function GetFolder(prevPath As String) As String
Dim strPath As String
    BrowseForFolder dhcCSIdlMyComputer, dhcBifReturnOnlyFileSystemDirs, strPath
    If Len(strPath) = 1 Then
        GetFolder = prevPath
    Else
        GetFolder = strPath
    End If
End Function

VERSION 5.00
Begin VB.Form frmTrayIcon 
   BorderStyle     =   0  'None
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   Icon            =   "frmClientTrayIcon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleWidth      =   885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "frmClientTrayIcon.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      Picture         =   "frmClientTrayIcon.frx":0B14
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   240
      Width           =   240
   End
End
Attribute VB_Name = "frmTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub ChangeIcon(iconnum As Long)
    Dim nid As NOTIFYICONDATA  ' icon information
    Dim retval As Long  ' return value
    
    Me.Icon = picIcon(iconnum).Picture
    
    ' Put the icon settings into the structure.
    With nid
      .cbSize = Len(nid)  ' size of structure
      .hWnd = Me.hWnd  ' owner of the icon and processor of its messages
      .uID = 1  ' unique identifier for the window's tray icons
      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP  ' provide icon, message, and tool tip text
      .uCallbackMessage = PK_TRAYICON  ' message to use for icon events
      .hIcon = Me.Icon    ' handle to the icon to actually display in the tray
      .szTip = "Yippee! Messenger (" & IIf(iconnum = 1, "Online", "Offline") & ")" & vbNullChar ' tool tip text for icon
    End With
    
    ' Add the icon to the system tray.
    retval = Shell_NotifyIcon(NIM_MODIFY, nid)
End Sub

Private Sub Form_Load()
    Dim nid As NOTIFYICONDATA  ' icon information
    Dim retval As Long  ' return value
    
    ' Put the icon settings into the structure.
    With nid
      .cbSize = Len(nid)  ' size of structure
      .hWnd = Me.hWnd  ' owner of the icon and processor of its messages
      .uID = 1  ' unique identifier for the window's tray icons
      .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP  ' provide icon, message, and tool tip text
      .uCallbackMessage = PK_TRAYICON  ' message to use for icon events
      .hIcon = Me.Icon  ' handle to the icon to actually display in the tray
      .szTip = "Yippee! Messenger (Offline)" & vbNullChar  ' tool tip text for icon
    End With
    
    ' Add the icon to the system tray.
    retval = Shell_NotifyIcon(NIM_ADD, nid)
    ' Set the new window procedure for window Form1.
    pOldProc = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim nid As NOTIFYICONDATA  ' icon information
Dim retval As Long  ' return value
    
    ' Load the structure with just the identifying information.
    With nid
        .cbSize = Len(nid)  ' size of structure
        .hWnd = Me.hWnd  ' handle of owning window
        .uID = 1  ' unique identifier
    End With
    retval = Shell_NotifyIcon(NIM_DELETE, nid)
    
    ' Make the old window procedure the current window procedure.
    retval = SetWindowLong(Me.hWnd, GWL_WNDPROC, pOldProc)
End Sub


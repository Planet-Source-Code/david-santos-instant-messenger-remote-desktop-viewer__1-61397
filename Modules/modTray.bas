Attribute VB_Name = "modTray"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2
Public Const WS_VISIBLE = &H10000000
Public Const GWL_STYLE = (-16)

Public Const WM_SYSCOMMAND = &H112

Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64  ' Windows 2000: make this String * 128
  ' The following data members are only valid in Windows 2000!
  ' (uncomment the following lines to use them)
  'dwState As Long
  'dwStateMask As Long
  'szInfo As String * 256
  'uTimeoutOrVersion As Long
  'szInfoTitle As String * 64
  'dwInfoFlags As Long
End Type

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_CLOSE = &H10

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const GWL_WNDPROC = (-4)

Public pOldProc As Long
Public Const PK_TRAYICON = &H401  ' program-defined message for tray icon action

Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    
    Case PK_TRAYICON
        
        Select Case lParam
        Case WM_RBUTTONUP
            'Popup the main form's File menu when we right-click
            frmMain.PopupMenu frmMain.mnuFile
        
        Case WM_LBUTTONDBLCLK
            'show the main window when we double-click
            frmMain.WindowState = vbNormal
            On Error Resume Next
            frmMain.Show
            If UserIndex = -1 Then frmLog.Show
            On Error GoTo 0
            frmMain.ZOrder 0
            If UserIndex = -1 Then frmLog.ZOrder 0

        End Select
        WindowProc = 1  ' this return value doesn't really matter
    
    Case Else
        ' Pass the message to the procedure Visual Basic provided.
        WindowProc = CallWindowProc(pOldProc, hWnd, uMsg, wParam, lParam)
    
    End Select
End Function


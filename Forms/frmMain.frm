VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Yippee! Messenger"
   ClientHeight    =   4350
   ClientLeft      =   12435
   ClientTop       =   1035
   ClientWidth     =   2850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   2850
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   2760
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   "onLine"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":102A
            Key             =   "offLine"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   4110
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   423
      Style           =   1
      SimpleText      =   "Disconnected"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2760
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.PictureBox picCover 
      BackColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4035
      ScaleWidth      =   2715
      TabIndex        =   2
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox picLogo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   600
         Picture         =   "frmMain.frx":157E
         ScaleHeight     =   1095
         ScaleWidth      =   1575
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblConnect 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sign in!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   600
         MouseIcon       =   "frmMain.frx":6FDC
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2160
         Width           =   1575
      End
   End
   Begin MSComctlLib.TreeView tvwFriends 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7223
      _Version        =   393217
      Indentation     =   388
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSignIn 
         Caption         =   "&Sign in"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "&Register"
      End
      Begin VB.Menu mnuBlank2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "&Settings"
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewOffline 
         Caption         =   "View Offline Messages"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBlank3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAddManual 
         Caption         =   "Add friend by ID"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "&Chat"
      End
      Begin VB.Menu mnuSend 
         Caption         =   "Send a &File"
      End
      Begin VB.Menu mnuAddFriend 
         Caption         =   "&Add as a friend"
      End
      Begin VB.Menu mnuRemoveFriend 
         Caption         =   "&Remove friend"
      End
      Begin VB.Menu mnuSendOffline 
         Caption         =   "Send offline &message"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'stores users list of friends from the server
Dim FriendList() As String

Dim ChatWindow As New Collection

Dim retryConnect As Integer

Dim scrdata() As Byte
Dim LastDIB As cDIBSection
Dim cDIB As cDIBSection

'used for taking a snapshot of the desktop
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest



Public Function NewTransfer(FileName As String, FileSize As Long, Remote As String, Port As Long, mode As FILETRANSFERMODE) As Boolean
Dim fTransfer As New frmFileTransfer

    Load fTransfer
    With fTransfer
        .FileName = FileName
        .FileSize = FileSize
        .Remote = Remote
        .mode = mode
        
        .Port = Port
        .StartTransfer
        'get new value of port if changed
        Port = .Port
        
        .Show , Me
    End With
    
    NewTransfer = True

End Function

Public Sub Connect()
On Error GoTo ErrConnect
    StatusBar1.SimpleText = "Connecting..."
    lblConnect.Caption = "Connecting to Yippee! as " & UserName
    
    Winsock1.Close
    
    'reserve port for winsock control
    Winsock1.Bind Winsock1.LocalPort
    
    'Send connection request to server
    Winsock1.SendData "CON:" & UserName & Chr(2) & Password
    Exit Sub

ErrConnect:
    If Err.Number <> 126 Then
        MsgBox "There was a problem connecting to the server:" & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Error Connecting"
        Disconnect
    End If
End Sub

Public Sub Disconnect()
Dim i As Integer
    If UserIndex > -1 Then
        'Send disconnection notice to server
        Winsock1.SendData "DIS:" & UserIndex
        Winsock1.Close
    End If
    
    UserIndex = -1
    tvwFriends.Nodes.Clear
    picCover.Visible = True
    mnuViewOffline.Enabled = False
    
    UserName = ""
    Password = ""
    
    StatusBar1.SimpleText = "Disconnected"
    lblConnect.Caption = "Sign in!"
    frmTrayIcon.ChangeIcon 0
    
    While ChatWindow.Count > 0
        Unload ChatWindow(1)
        ChatWindow.Remove 1
    Wend
    
    mnuSignIn.Caption = "Sign in"
End Sub


Private Sub Form_Load()
Dim sh As Long
Dim sw As Long

    'Prevent more than one instance of program from running
    If App.PrevInstance Then End
    
    Me.Show
    
    Me.Width = 3000
    Me.Left = Screen.Width - Me.Width - 150
    Me.Top = 500
    
    Load frmTrayIcon
    
    UserIndex = -1
    
    BufferFolder = GetSystemFolder(CSIDL_PERSONAL) & "\Buffer"
    DownloadsFolder = GetSystemFolder(CSIDL_PERSONAL) & "\Downloads"
    LogsFolder = GetSystemFolder(CSIDL_PERSONAL) & "\Logs"
    
    MakeSureDirectoryExists BufferFolder
    MakeSureDirectoryExists DownloadsFolder
    MakeSureDirectoryExists LogsFolder
    
    Winsock1.Protocol = sckUDPProtocol
    
    'Set the address of the server from the INI file
    Winsock1.RemoteHost = ReadINI(App.path & "\config.ini", "Settings", "RemoteHost", "169.192.0.1")
    'Set the port of the server from the INI file
    Winsock1.RemotePort = ReadINI(App.path & "\config.ini", "Settings", "RemotePort", "9004")
    Winsock1.LocalPort = ReadINI(App.path & "\config.ini", "Settings", "LocalPort", "9005")
   
    baseFileTransferPort = CInt(ReadINI(App.path & "\config.ini", "Settings", "BFTP", "9100"))
   
    mFontName = ReadINI(App.path & "\config.ini", "Chat", "Font", "Arial")
    mColor = CLng(ReadINI(App.path & "\config.ini", "Chat", "Color", CStr(RGB(0, 0, 255))))
    mFontSize = CInt(ReadINI(App.path & "\config.ini", "Chat", "FontSize", "10"))
    
    frmLog.Show vbModal
    
    sh = Screen.Height / Screen.TwipsPerPixelY
    sw = Screen.Width / Screen.TwipsPerPixelX
    
    
    Set LastDIB = New cDIBSection
    Set cDIB = New cDIBSection
    
    cDIB.Create sw, sh
    LastDIB.Create sw, sh
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Me.Hide
        Cancel = 1
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Exit Sub
    End If
    If Me.Height < 1000 Then Me.Height = 1000
    'If Me.Height > 8000 Then Me.Height = 8000
    'If Me.Width > 3000 Then Me.Width = 3000
    'If Me.Width < 3000 Then Me.Width = 3000
    
    tvwFriends.Width = Me.Width - 8 * 15
    tvwFriends.Height = Me.Height - StatusBar1.Height - 54 * 15
    
    picCover.Height = tvwFriends.Height
    picCover.Width = tvwFriends.Width
    
    lblConnect.Left = picCover.Width / 2 - lblConnect.Width / 2
    lblConnect.Top = picCover.Height / 2 - lblConnect.Height / 2 + picLogo.Height / 2
    picLogo.Left = lblConnect.Left
    picLogo.Top = lblConnect.Top - picLogo.Height - 8 * 15
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Disconnect
    Unload frmTrayIcon
End Sub

Private Sub SignIn()
    If UserIndex = -1 Then
        On Error Resume Next
        frmLog.Show vbModal
        frmLog.ZOrder 0
    Else
        If MsgBox("Are you sure you want to disconnect?", vbYesNo + vbQuestion, "Disconnect from server") = vbYes Then
            Disconnect
        End If
    End If
End Sub

Private Sub lblConnect_Click()
    SignIn
End Sub

Private Sub lblConnect_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblConnect.ForeColor = RGB(255, 0, 0)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuAddFriend_Click()
    Winsock1.SendData "ADF:" & UserIndex & Chr(2) & tvwFriends.SelectedItem.Text
End Sub

Private Sub mnuAddManual_Click()
    frmAdd.Show vbModal
End Sub

Private Sub mnuChat_Click()
    tvwFriends_DblClick
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuRegister_Click()
    frmRegistration.Show vbModal
End Sub

Private Sub mnuRemoveFriend_Click()
    Winsock1.SendData "RMF:" & UserIndex & Chr(2) & tvwFriends.SelectedItem.Text
End Sub

Private Sub mnuSend_Click()
Dim i As Integer
Dim newChat As frmChat
    
    If Left(tvwFriends.SelectedItem.key, 4) = "user" Or Left(tvwFriends.SelectedItem.key, 6) = "friend" Then
        If tvwFriends.SelectedItem.Image = 1 Then
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = tvwFriends.SelectedItem.Text Then
                    ChatWindow.SendFile
                    
                    ChatWindow(i).Show
                    Exit Sub
                End If
            Next
            
            Set newChat = New frmChat
            
            newChat.Caption = tvwFriends.SelectedItem.Text
            newChat.Tag = UserName
            newChat.Show
            newChat.SendFile
            ChatWindow.Add newChat
        End If
    End If
End Sub

Private Sub mnuSendOffline_Click()
    frmSendOffline.txtTo.Text = tvwFriends.SelectedItem.Text
    frmSendOffline.Show vbModal
End Sub

Private Sub mnuSetup_Click()
    frmSetup.Show vbModal
End Sub

Private Sub mnuSignIn_Click()
    SignIn
End Sub

Private Sub mnuViewOffline_Click()
    Winsock1.SendData "CMD:OLM:" & UserIndex & Chr(2) & "ALL"
End Sub

Private Sub picCover_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblConnect.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblConnect.ForeColor = RGB(0, 0, 255)
End Sub

Private Sub tvwFriends_DblClick()
Dim i As Integer
Dim newChat As frmChat

    If Left(tvwFriends.SelectedItem.key, 4) = "user" Or Left(tvwFriends.SelectedItem.key, 6) = "friend" Then
        If tvwFriends.SelectedItem.Image = 1 Then
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = tvwFriends.SelectedItem.Text Then
                    ChatWindow(i).Show
                    Exit Sub
                End If
            Next
            
            Set newChat = New frmChat
            
            newChat.Caption = tvwFriends.SelectedItem.Text
            newChat.Tag = UserName
            newChat.Show
            ChatWindow.Add newChat
        End If
    End If

End Sub

Private Sub tvwFriends_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    If KeyCode = vbKeyReturn Then
    
        Dim newChat As frmChat
        
            If Left(tvwFriends.SelectedItem.key, 4) = "user" Or Left(tvwFriends.SelectedItem.key, 6) = "friend" Then
                If tvwFriends.SelectedItem.Image = 1 Then
                    For i = 1 To ChatWindow.Count
                        If ChatWindow(i).Caption = tvwFriends.SelectedItem.Text Then
                            ChatWindow(i).Show
                            Exit Sub
                        End If
                    Next
                    
                    Set newChat = New frmChat
                    
                    newChat.Caption = tvwFriends.SelectedItem.Text
                    newChat.Tag = UserName
                    newChat.Show
                    ChatWindow.Add newChat
                End If
            End If
    
    
    End If
End Sub

Private Sub tvwFriends_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim isOnline As Boolean
    If Button = 2 Then
    
        If tvwFriends.SelectedItem Is Nothing Then Exit Sub
        isOnline = (tvwFriends.SelectedItem.Image = 1)
        
        If Left(tvwFriends.SelectedItem.key, 6) = "friend" Then
            mnuChat.Enabled = isOnline
            mnuSend.Enabled = isOnline
            mnuAddFriend.Enabled = False
            mnuRemoveFriend.Enabled = True
            mnuSendOffline = Not isOnline
        ElseIf Left(tvwFriends.SelectedItem.key, 4) = "user" Then
            mnuChat.Enabled = isOnline
            mnuSend.Enabled = isOnline
            mnuAddFriend.Enabled = True
            mnuRemoveFriend.Enabled = False
            mnuSendOffline = Not isOnline
        Else
            mnuChat.Enabled = False
            mnuSend.Enabled = False
            mnuAddFriend.Enabled = False
            mnuRemoveFriend.Enabled = False
            mnuSendOffline = False
        End If
        
        PopupMenu mnuMenu
    End If
End Sub

Private Sub RemoveHeader(packet As String)
    packet = Right(packet, Len(packet) - 4)
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Dim UserList() As String
Dim newChat As frmChat
Dim bfound As Boolean
Dim curNode As Node
Dim DataArray() As String
Dim i As Integer
Dim fname As String


    If bytesTotal < 3 Then Exit Sub
    Winsock1.GetData data
    
    Select Case Left(data, 3)
    Case "STA"
        'Status packet
        RemoveHeader data

        Select Case Left(data, 3)
        Case "OFL"
            'Server shut down normally
            UserIndex = -1
            Disconnect
            
        Case "ONL"
            RemoveHeader data
            
            'jus in case we get more than 1 ONL message...
            If UserIndex > -1 Then Exit Sub
            
            'Server acknowledges user is online
            'so enable stuff here
            tvwFriends.Nodes.Add , , "Online", "Online", 1
            tvwFriends.Nodes.Add , , "Friends", "Friends"
            tvwFriends.Nodes.Item("Online").Expanded = True
            tvwFriends.Nodes.Item("Friends").Expanded = True
            tvwFriends.Nodes.Item("Online").Sorted = True
            tvwFriends.Nodes.Item("Friends").Sorted = True
            
            picCover.Visible = False
            StatusBar1.SimpleText = "Connected"
            mnuSignIn.Caption = "Sign out"
            mnuViewOffline.Enabled = True
            
            'set tray icon to Online mode
            frmTrayIcon.ChangeIcon 1
            
            UserIndex = data
            
            'requext friend list
            Winsock1.SendData "CMD:FRN:" & UserIndex
        
        Case "FRN"
            'Friends list packet
            RemoveHeader data
            
            FriendList = Split(data, Chr(2))
            
            Set curNode = tvwFriends.Nodes("Friends").Child
            
            'remove old friends list
            While Not (curNode Is Nothing)
                tvwFriends.Nodes.Remove curNode.Index
                Set curNode = tvwFriends.Nodes("Friends").Child
            Wend
            
            For i = 0 To UBound(FriendList)
                tvwFriends.Nodes.Add "Friends", tvwChild, "friend" & i, FriendList(i), 2
            Next
    
            'request online-user list
            Winsock1.SendData "CMD:LST:" & UserIndex
       
        Case "LST"
            'User list packet
            RemoveHeader data
    
            UserList = Split(data, Chr(2))
            
            Set curNode = tvwFriends.Nodes("Online").Child
            
            While Not (curNode Is Nothing)
                tvwFriends.Nodes.Remove curNode.Index
                Set curNode = tvwFriends.Nodes("Online").Child
            Wend
            
            For i = 0 To UBound(UserList)
                If FindFriend(UserList(i), curNode) Then
                    curNode.Image = 1
                Else
                    tvwFriends.Nodes.Add "Online", tvwChild, "user" & tvwFriends.Nodes("Online").Children, UserList(i), 1
                End If
            Next
            
            'request offline messages
            Winsock1.SendData "CMD:OLM:" & UserIndex & Chr(2) & "NEW"
            
        Case "OLM"
            RemoveHeader data
            
            DataArray = Split(data, Chr(2))
            
            'add message to Offline messages window
            With frmOfflineMsg.lvwOfflineMsgs.ListItems
                .Add , , DataArray(0)
                .Item(.Count).SubItems(1) = Format(DataArray(1), "mmm d, yy hh:mm AMPM")
                .Item(.Count).SubItems(2) = DataArray(2)
                .Item(.Count).Tag = DataArray(3)
                .Item(.Count).Bold = (DataArray(4) = "F")
                .Item(.Count).ListSubItems(1).Bold = .Item(.Count).Bold
                .Item(.Count).ListSubItems(2).Bold = .Item(.Count).Bold
                frmOfflineMsg.StatusBar1.SimpleText = .Count & " offline message(s)"
                frmOfflineMsg.Caption = .Count & " Offline Message(s)"
            End With

            frmOfflineMsg.Show , Me
        
        End Select
        
    Case "USR"
        RemoveHeader data
        
        Select Case Left(data, 3)
        Case "FIL"
            'File Transfer Request
            RemoveHeader data
            DataArray = Split(data, Chr(2))
            
            'create a new Accept Deny form for each file request
            Dim fAccept As New frmAcceptDeny

            'Transfer data to the form
            fAccept.LoadArray DataArray
            
            fAccept.Show , Me
            
        Case "JOI"
            'JOIN packet
            RemoveHeader data
            
            'Check if user is in friends list
            If FindFriend(data, curNode) Then
                'set friend's icon to online
                curNode.Image = 1
            Else
                'add the user to the online user's list
                tvwFriends.Nodes.Add "Online", tvwChild, "user" & data, data, 1
            End If
            
        Case "PRT"
            'PART packet
            RemoveHeader data
            
            'Check if user is in friends list
            If FindFriend(data, curNode) Then
                'set friend's icon to offline
                curNode.Image = 2
            Else
                'find the users anme in the online list and remove it
                Set curNode = tvwFriends.Nodes("Online").Child
                Do While Not (curNode Is Nothing)
                    If curNode.Text = data Then
                        tvwFriends.Nodes.Remove curNode.Index
                        Exit Do
                    End If
                    Set curNode = curNode.Next
                Loop
            End If
    
            'find the chat window with the user's name
            'and, if found, notify that he/she has left
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = data Then
                    ChatWindow(i).AddText "*** " & data & " has left...", RGB(0, 255, 0), , 12
                    Exit For
                End If
            Next
            
        Case "BUZ"
            'Message packet
            RemoveHeader data
    
            DataArray = Split(data, Chr(2))
    
            'find the chat window with the user's name
            'and, if found, BUZZ it!
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = DataArray(0) Then
                    If Not ChatWindow(i).Buzzing Then
                        PlayWav App.path & "\Media\doorbell.wav"
                        ChatWindow(i).AddText "***BUZZ!", RGB(255, 0, 0), , 12
                        ChatWindow(i).Buzz
                        bfound = True
                    Else
                        bfound = True
                    End If
                End If
            Next
            
            'not found, so create a new window and... BUZZ it!
            If Not bfound Then
                Set newChat = New frmChat
                newChat.Caption = DataArray(0)
                newChat.Tag = UserName
                newChat.Show
                ChatWindow.Add newChat
                newChat.AddText "***BUZZ!", RGB(255, 0, 0), , 12
                newChat.Buzz
            End If
            
        Case "MSG"
            'Message packet
            RemoveHeader data
    
            DataArray = Split(data, Chr(2))
    
            'find the chat window with the user's name
            'and, if found, add message
                        
            
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = DataArray(0) Then
                    
                    If Left(DataArray(0), 1) = "@" And DataArray(0) <> "@Server" Then
                        DataArray(0) = ""
                    Else
                        DataArray(0) = DataArray(0) & "> "
                    End If
                    
                    ChatWindow(i).AddText DataArray(0) & DataArray(1), CLng(DataArray(2)), DataArray(3), CInt(DataArray(4))
                    bfound = True
                End If
            Next
            
            'not found, so create a new window and add message
            If Not bfound Then
                Set newChat = New frmChat
                
                newChat.Caption = DataArray(0)
                newChat.Tag = UserName
                newChat.Show
                
                ChatWindow.Add newChat
                    
                If Left(DataArray(0), 1) = "@" Then
                    DataArray(0) = ""
                    newChat.mnuSendFile.Enabled = False
                Else
                    DataArray(0) = DataArray(0) & "> "
                End If
                
                newChat.AddText DataArray(0) & DataArray(1), CLng(DataArray(2)), DataArray(3), CInt(DataArray(4))
            End If
        
        End Select
        
    Case "PIC"
        'Picture data packet
        'we may want to replace this with a direct file transfer via TCP
        Dim tempname As String
        Dim namelen As Long
        Dim partno As Long
        Dim partof As Long
        RemoveHeader data
        
        Select Case Left(data, 3)
            
        Case "GOT"
            Dim picname As String
            Dim lastpart As Integer
            RemoveHeader data
            
            lastpart = Split(data, Chr(2))(0)
            picname = Split(data, Chr(2))(1)
            SendPic picname, lastpart + 1
        
        Case "SND"
            RemoveHeader data
            
            partno = Val(Left(data, 3))
            data = Right(data, Len(data) - 3)
            
            partof = Val(Left(data, 3))
            data = Right(data, Len(data) - 3)
            
            namelen = Val(Left(data, 3))
            data = Right(data, Len(data) - 3)
            
            tempname = Left(data, namelen)
            data = Right(data, Len(data) - namelen - 1)
            
            
            Dim targetfile As String
            
            MakeSureDirectoryExists BufferFolder
            
            targetfile = BufferFolder & "\" & tempname & ".jpg"
            
            If partno = 0 Then
                If Dir(targetfile) <> "" Then Kill targetfile
            End If
                    
            Open targetfile For Binary As 1
                Put #1, (partno * 7168) + 1, data
            Close
            
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = tempname Then
                    If partno = 0 Then ChatWindow(i).pgbDownload.Visible = True
                    ChatWindow(i).pgbDownload.Max = partof
                    ChatWindow(i).pgbDownload.Value = partno
                    Exit For
                End If
            Next
            
            If partno = partof Then
                ChatWindow(i).pgbDownload.Visible = False
                
                For i = 1 To ChatWindow.Count
                    If ChatWindow(i).Caption = tempname Then
                        ChatWindow(i).Picture1.Picture = LoadPicture(targetfile)
                        ChatWindow(i).Picture1.Refresh
                        Exit For
                    End If
                Next
            Else
                Sleep 10
                Winsock1.SendData "PIC:GOT" & Chr(2) & partno & Chr(2) & UserIndex & Chr(2) & tempname
            End If
        
        End Select
        
    Case "REG"
        'Registration packet
        RemoveHeader data
        DataArray = Split(data, Chr(2))
    
        With frmRegistration
            'enable the registration form
            .Connected
            'retrieve user info
            UserName = DataArray(0)
            'fail gracefully
            fname = BufferFolder & "\" & UserName & ".jpg"
            If Dir(fname) <> "" Then .picMyPic.Picture = LoadPicture(fname)
            .txtName.Text = DataArray(1)
            .txteMail.Text = DataArray(2)
            .txtOther.Text = DataArray(3)
        End With
            
        
    Case "NTF"
        'Notify packet
        RemoveHeader data
        MsgBox data, vbInformation, "Server Notification"
    
    Case "ERR"
        'Error packet
        RemoveHeader data
        MsgBox data, vbInformation, "Error"
        Disconnect
    
    Case "ACK"
        Winsock1.SendData "ACK!" & UserIndex

    Case "ADF"
        RemoveHeader data
        
        Select Case data
        Case "OK"
            Unload frmAdd
        Case "FAIL"
            MsgBox "No such account was found", vbInformation, "Add Friend"
        End Select
        
    Case "MON"
        'request screenshot update
        Winsock3.Close
        Winsock3.Bind 9200
        Winsock3.Listen
        
    Case Else
        'Unknown packet...
        Debug.Print data
    
    End Select

End Sub

Public Function SendPic(FileName As String, part As Long)
Dim picdata As String, partdata As String, numparts As Long
    If Dir(FileName) <> "" Then
        Open FileName For Binary As 1
        picdata = String(LOF(1), Chr(0))
        Get #1, , picdata
        Close
        
        numparts = Len(picdata) \ 7168
        
        If Len(picdata) Mod 7168 > 0 Then
            numparts = numparts + 1
        End If
        
        partdata = String(7168, Chr(0))
        
        partdata = Mid(picdata, part * 7168 + 1, 7168)
        Winsock1.SendData "PIC:UPP:" & Format(part, "000") & Format(numparts, "000") & Format(Len(FileName), "000") & FileName & Chr(2) & partdata
        
        picdata = ""
    End If
End Function

Private Function FindFriend(name As String, curNode As Node) As Boolean
Dim bfound As Boolean
    Set curNode = tvwFriends.Nodes("Friends").Child
    
    bfound = False
    
    Do While Not (curNode Is Nothing)
        If name = curNode.Text Then
            bfound = True
            Exit Do
        End If
        Set curNode = curNode.Next
    Loop
    
    FindFriend = bfound
End Function

Public Function RemoveChatWindow(WindowName As String)
Dim i As Integer
    For i = 1 To ChatWindow.Count
        If ChatWindow(i).Caption = WindowName Then
            ChatWindow.Remove i
            Exit For
        End If
    Next
End Function

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
    Winsock3.Close
    Winsock3.Accept requestID
    SendScreen
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
Dim data As String
    Winsock3.GetData data
    If data = "UPDATE" Then
        SendScreen
    End If
End Sub


Private Sub SendScreen()
Dim hDesktopDC As Long
Dim hWndDesktop As Long
Dim sw As Long, sh As Long
Dim requestID As Long

    
    sh = Screen.Height / Screen.TwipsPerPixelY
    sw = Screen.Width / Screen.TwipsPerPixelX
    
    
    hWndDesktop = GetDesktopWindow
    hDesktopDC = GetDC(hWndDesktop)
    
    
    ReDim scrdata(sw * sh * 3)
    ' Try to create a Device Independent Bitmap (DIB)
    
    BitBlt LastDIB.hDC, 0, 0, sw, sh, cDIB.hDC, 0, 0, SRCCOPY
    BitBlt cDIB.hDC, 0, 0, sw, sh, hDesktopDC, 0, 0, SRCERASE
    BitBlt LastDIB.hDC, 0, 0, sw, sh, hDesktopDC, 0, 0, SRCCOPY
    ReleaseDC hWndDesktop, hDesktopDC
    
    ' Using the DIB, we can convert picture data
    ' into a byte array that we can send over the network
    
    cDIB.GetByteArray scrdata
    Dim cComp As New cCompression
    cComp.CompressByteArray scrdata, COMPRESS_BEST_SPEED
    Set cComp = Nothing
    Winsock1.SendData "SCR:SET:" & sw & Chr(2) & sh & Chr(2) & UBound(scrdata)
    Sleep 5
    Winsock3.SendData scrdata
    Erase scrdata
    
    'Else
        'Couldn't create a DIB, so send a blank screen
    '    ReDim data(0)
    '    Winsock3.SendData data
    '    Erase data
    'End If


    'Need to release the DC we obtained, in order
    'to free the resources we used
End Sub

Private Sub Winsock3_SendComplete()
    Winsock3.Close
End Sub

VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm mdiServer 
   BackColor       =   &H8000000C&
   Caption         =   "Yippee! IM Server"
   ClientHeight    =   8220
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11595
   Icon            =   "mdiServer.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picUsers 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8220
      Left            =   9120
      ScaleHeight     =   8220
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   0
      Width           =   2475
      Begin VB.ListBox lstUsers 
         Height          =   8220
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Timer tmrAck 
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuMultiSend 
         Caption         =   "&Send a File"
      End
      Begin VB.Menu mnuManage 
         Caption         =   "Manage &Users"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuTileH 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuTileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuMonitor 
         Caption         =   "&Monitor"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuABout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "mdiServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim maxUsers As Long
Dim currentmonitor As Long

Dim ChatWindow As New Collection

Private Sub lstUsers_DblClick()
Dim newChat As frmChat
Dim i As Integer

    For i = 1 To ChatWindow.Count
        If ChatWindow(i).Caption = lstUsers.List(lstUsers.ListIndex) Then
            ChatWindow(i).Show
            Exit Sub
        End If
    Next
    
    Set newChat = New frmChat
    
    newChat.Caption = lstUsers.List(lstUsers.ListIndex)
    newChat.Show
    ChatWindow.Add newChat

End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub MDIForm_Load()
    'read settings from the SERVER.INI file
    Winsock1.RemoteHost = ReadINI(App.Path & "\server.ini", "Settings", "RemoteHost", Winsock1.LocalIP)
    Winsock1.RemotePort = ReadINI(App.Path & "\server.ini", "Settings", "RemotePort", "9004")
    Winsock1.LocalPort = ReadINI(App.Path & "\server.ini", "Settings", "LocalPort", "9004")
    maxUsers = ReadINI(App.Path & "\server.ini", "Settings", "MaxUsers", "150")
    
    mFontName = ReadINI(App.Path & "\server.ini", "Chat", "Font", "Arial")
    mColor = CLng(ReadINI(App.Path & "\server.ini", "Chat", "Color", CStr(RGB(0, 0, 255))))
    mFontSize = CInt(ReadINI(App.Path & "\server.ini", "Chat", "FontSize", "10"))
    
    Me.Show
    frmSplash.Show , Me
    'maxUsers is the number of allowable users on the system
    'limit the value between 10 and 999
    If maxUsers < 10 Then maxUsers = 10
    If maxUsers > 999 Then maxUsers = 999
    
    'reserve memory for the user records
    ReDim User(maxUsers - 1)
    
    'create a blank ban list with 1 entry
    
    'set up the winsock control to use UDP
    Winsock1.Protocol = sckUDPProtocol
    
    'bind control to the local port set above in the INI file
    'winsock will use this port to catch inbound packets
    Winsock1.Bind Winsock1.LocalPort
    
    'open the user database
    OpenDB App.Path & "\userdb.mdb"
        
    'create a new recordset
    Set UserRS = New ADODB.Recordset
    
    
    
    Load frmStatus
    
    'show the server status
    AddStatus "Yippee! Server v0.31 initializing..."
    AddStatus "Server IP: " & Winsock1.LocalIP
    AddStatus "Server port: " & Winsock1.LocalPort
    AddStatus "MaxUsers: " & maxUsers
    AddStatus IIf(Winsock1.State = sckOpen, "Ready!...", "Socket not open!") & vbCrLf
    
    'Register the server as a user on the system
    'so that users can send messages/files to the server
    User(0).UserName = "@Server"
    'set the IP address so that system recognizes that this user record is being used
    User(0).IPAddress = Winsock1.LocalIP
    
    User(1).UserName = "@Everyone"
    'set the IP address so that system recognizes that this user record is being used
    User(1).IPAddress = Winsock1.LocalIP
    
    
    lstUsers.AddItem "@Everyone"
End Sub


'sends data to a specific user
'IP add and port are taken from the user's record
Private Sub SendtoUser(User As UserRecord, data As String)
On Error GoTo errSend
    Winsock1.RemoteHost = User.IPAddress
    Winsock1.RemotePort = User.Port
    Winsock1.SendData data
    Exit Sub
errSend:
    If Err.Number = 10013 Then Debug.Print "Socket blocked!"
End Sub

Public Sub SendtoName(name As String, data As String)
On Error GoTo errSend
Dim i As Integer

    If name = "@Everyone" Then
        For i = 2 To lastUser
            Winsock1.RemoteHost = User(i).IPAddress
            Winsock1.RemotePort = User(i).Port
            Winsock1.SendData data
        Next
        Exit Sub
    End If

    For i = 2 To lastUser
        If User(i).IPAddress <> "" Then
            If User(i).UserName = name Then
                Winsock1.RemoteHost = User(i).IPAddress
                Winsock1.RemotePort = User(i).Port
                Winsock1.SendData data
                Exit For
            End If
        End If
    Next
    
    Exit Sub
errSend:
    If Err.Number = 10013 Then Debug.Print "Socket blocked!"
End Sub

Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuWindow
    End If
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 780 Then Exit Sub
    lstUsers.Height = Me.Height - 58 * 15
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim i As Integer
Dim OLU As Long

    For i = 0 To lstUsers.ListCount - 1
        If Not (Left(lstUsers.List(i), 1) = "@") Then OLU = OLU + 1
    Next
    
    
    If OLU > 0 Then
        If MsgBox("There are still " & OLU & " user(s) connected to the server. Are you sure you want to quit?", vbQuestion + vbYesNo + vbDefaultButton2, "Closing Server") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    
    For i = 2 To lastUser
        If User(i).IPAddress <> "" Then
            SendtoUser User(i), "STA:OFL"
        End If
    Next
    
    Unload frmStatus
    Winsock1.Close
    Set UserRS = Nothing
    CloseDB
End Sub

Private Sub RemoveHeader(packet As String)
    'remove header (first 4 chars)
    packet = Right(packet, Len(packet) - 4)
End Sub

Private Sub mnuABout_Click()
    frmAboutServer.Show vbModal
End Sub

Private Sub mnuArrange_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Public Sub MonitorUser()
    SendtoUser User(currentmonitor), "MON"
    frmMonitor.Winsock1.Close
    frmMonitor.Winsock1.RemoteHost = User(currentmonitor).IPAddress
    frmMonitor.Winsock1.RemotePort = 9200
    frmMonitor.Caption = "Monitoring " & User(currentmonitor).UserName
End Sub

Private Sub mnuManage_Click()
    frmDB.Show
End Sub

Private Sub mnuMonitor_Click()
Dim i As Integer
    For i = 2 To lastUser
        If lstUsers.List(lstUsers.ListIndex) = User(i).UserName Then
            currentmonitor = i
            frmMonitor.Show
            MonitorUser
            Exit For
        End If
    Next
End Sub

Private Sub mnuMultiSend_Click()
    frmSendMulti.Show
End Sub

Private Sub mnuTileH_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileV_Click()
    Me.Arrange vbTileVertical
End Sub

' this provides a way of checking if the user is still online.
' every 5 seconds (or set according to the timer's Interval property)
' the system decrements each online user's timeout value, which is initially set to
' 60 (seconds). When the user's timer reaches 30, the server sends an ACK? request
' to the client. If the client does not respond with an "ACK!" in the next 30 seconds,
' the user is dropped from the list.
Private Sub tmrAck_Timer()
Dim i As Integer
Dim j As Integer
    For i = 2 To lastUser
        If User(i).IPAddress <> "" Then
            'decrement user's timeout
            User(i).Timeout = User(i).Timeout - tmrAck.Interval / 1000
            'time to send an ACK? for this user
            If User(i).Timeout = 30 Then SendtoUser User(i), "ACK?"
            
            'user has timed out
            If User(i).Timeout = 0 Then
                
                'notify all users
                For j = 2 To lastUser
                    If User(j).IPAddress <> "" Then
                        SendtoUser User(j), "USR:PRT:" & User(i).UserName
                    End If
                Next
                
                'remove user from list
                For j = 0 To lstUsers.ListCount
                    If User(i).UserName = lstUsers.List(j) Then
                        lstUsers.RemoveItem j
                        Exit For
                    End If
                Next
                
                'remove user from record
                With User(i)
                    .IPAddress = ""
                    .UserID = 0
                    .Port = ""
                    .UserName = ""
                End With
            
            End If
        End If
    Next
End Sub

'the core of system
'this event is triggered when data is sent to the local IP address on
'the bound port
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim tempRS As New ADODB.Recordset
Dim n As Integer
Dim data As String
Dim DataArray() As String
Dim WhoIsOnline As String
Dim GetIP As String, GetPort As String, GetName As String
Dim RequesterID As Long, FriendID As Long, Requester As String
Dim picData As String, PartData As String, NumParts As Long
Dim i As Integer, j As Integer
Dim bFound As Boolean
Dim newChat As frmChat
                
    If bytesTotal = 1 Then Exit Sub
    
    Winsock1.GetData data
    Winsock1.RemoteHost = Winsock1.RemoteHostIP
    
    Select Case Left(data, 3)
    Case "SCR"
        RemoveHeader data
        
        Select Case Left(data, 3)
        Case "SET"
            RemoveHeader data
            DataArray = Split(data, Chr(2))
            frmMonitor.SetDIBSize CLng(DataArray(0)), CLng(DataArray(1)), CLng(DataArray(2))
        
        Case "OK!"
            frmMonitor.UpdateDisplay
        
        End Select
        
    Case "FIL"
        RemoveHeader data
        Select Case Left(data, 3)
        Case "REQ"
            RemoveHeader data
                
            DataArray = Split(data, Chr(2))
            
            'look for user with corresponding name
            For i = 2 To lastUser
                If User(i).UserName = DataArray(1) Then
                    SendtoUser User(i), "USR:FIL:" & DataArray(0) & Chr(2) & DataArray(2) & Chr(2) & DataArray(3) & Chr(2) & DataArray(4) & Chr(2) & DataArray(5) & Chr(2) & DataArray(6)
                    Exit For
                 End If
            Next
            
        End Select
        
    Case "ACK"      'Acknowledge
        RemoveHeader data
        
        'Client has acknowledged, reset user's timeout
        User(data).Timeout = 120
    
    Case "SOL"      'Send offline message
        RemoveHeader data
        
        DataArray = Split(data, Chr(2))
        Requester = User(DataArray(0)).UserName
        RequesterID = User(DataArray(0)).UserID
        
        Set UserRS = QueryDB("SELECT UserName, UserID From RegUsers")
        'get friends's database ID
        UserRS.MoveFirst
        UserRS.Find "UserName = '" & DataArray(1) & "'"
        FriendID = UserRS("UserID")
        
        UserRS.Close
    
        'write offline message in database
        Set UserRS = QueryDB("SELECT * From OfflineMsgs")
        UserRS.AddNew
        UserRS("Sender") = RequesterID
        UserRS("Recipient") = FriendID
        UserRS("Timestamp") = Date & " " & Time
        UserRS("Message") = DataArray(2)
        UserRS.Update
        UserRS.Close
    
    'Add a Friend
    Case "ADF"
        RemoveHeader data
        
        DataArray = Split(data, Chr(2))
        Requester = User(DataArray(0)).UserName
        'get user's registered ID
        RequesterID = User(DataArray(0)).UserID

        'get friends's registered ID with the UserNo
        Set UserRS = QueryDB("SELECT * From RegUsers WHERE UserNo = """ & DataArray(1) & """")
        If UserRS.RecordCount > 0 Then
            FriendID = UserRS("UserID")
            UserRS.Close
        Else
            'UserNo not found, get friends's registered ID with the UserName
            Set UserRS = QueryDB("SELECT * From RegUsers WHERE UserName = """ & DataArray(1) & """")
            If UserRS.RecordCount = 0 Then
                'Username not found either
                SendtoUser User(Split(data, Chr(2))(0)), "ADF:FAIL"
                Exit Sub
            Else
                FriendID = UserRS("UserID")
                UserRS.Close
            End If
        End If

        'add friend to database
        Set UserRS = QueryDB("SELECT * From FriendsList")
        UserRS.AddNew
        UserRS("UserID") = RequesterID
        UserRS("FriendID") = FriendID
        UserRS.Update
        UserRS.Close

        'update user's friend list
        Winsock1.SendData "STA:FRN:" & MakeFriendList(RequesterID)
        Winsock1.SendData "ADF:OK"

    Case "RMF"
        'ReMove Friend
        RemoveHeader data
        
        DataArray = Split(data, Chr(2))
        
        RequesterID = User(DataArray(0)).UserID

        'get friends's registered ID
        Set UserRS = QueryDB("SELECT * From RegUsers WHERE UserName = """ & DataArray(1) & """")
        FriendID = UserRS("UserID")
        UserRS.Close

        Set UserRS = QueryDB("SELECT * From FriendsList WHERE UserID=" & RequesterID & " AND FriendID=" & FriendID & "")
        UserRS.Delete
        UserRS.Update
        UserRS.Close

        Winsock1.SendData "STA:FRN:" & MakeFriendList(RequesterID)
    
    Case "REG"
        'REGister
        RemoveHeader data
        
        DataArray = Split(data, Chr(2))
        
        Set UserRS = QueryDB("SELECT * FROM RegUsers WHERE UserNo='" & DataArray(1) & "'")
        
        If UserRS.RecordCount = 0 Then
            Winsock1.SendData "ERR:The ID number """ & DataArray(1) & """ was not found."
            Exit Sub
        End If
        
        If UserRS("Password") <> DataArray(2) Then
            Winsock1.SendData "ERR:The password you typed was incorrect."
            Exit Sub
        End If
        
        If IsNull(UserRS("UserName")) Then
            If DataArray(0) = "" Then
                Winsock1.SendData "NTF:You haven't selected a login name yet. Please enter a name in the box provided."
                Exit Sub
            Else
                If Left(DataArray(0), 1) = "@" Then
                    Winsock1.SendData "NTF:You cannot use a name starting with ""@""."
                    Exit Sub
                End If
                
                Set tempRS = QueryDB("SELECT * FROM RegUsers WHERE UserName='" & DataArray(0) & "'")
                If tempRS.RecordCount > 0 Then
                    Winsock1.SendData "NTF:The name """ & DataArray(0) & """ has already been taken. Please try another name."
                    tempRS.Close
                    Exit Sub
                End If
                
                UserRS("UserName") = DataArray(0)
                UserRS.Update
                
                Winsock1.SendData "NTF:You have been registered into the system as """ & DataArray(0) & """."
            End If
        End If
        
        Winsock1.SendData "REG:" & UserRS("UserName") & Chr(2) & UserRS("FullName") & Chr(2) & UserRS("Email") & Chr(2) & UserRS("Other")
        
        UserRS.Close
    
    Case "UPD"
        'UPDate user's info
        DataArray = Split(data, Chr(2))
        
        Set UserRS = QueryDB("SELECT * FROM RegUsers WHERE UserNo='" & DataArray(1) & "'")
        
        If UserRS.RecordCount = 0 Then
            Winsock1.SendData "ERR:The ID number """ & DataArray(1) & """ was not found."
            Exit Sub
        End If
        
        UserRS("FullName") = IIf(DataArray(3) = "", Null, DataArray(3))
        UserRS("Email") = IIf(DataArray(4) = "", Null, DataArray(4))
        UserRS("Other") = IIf(DataArray(5) = "", Null, DataArray(5))
        
        If DataArray(6) <> "" Then UserRS("Password") = DataArray(6)

        UserRS.Update
        UserRS.Close
    
        Winsock1.SendData "NTF:The system has been updated."
    
    Case "CON"
        'Connection request
        RemoveHeader data
        
        DataArray = Split(data, Chr(2))
        
        GetIP = Winsock1.RemoteHostIP
        GetPort = Winsock1.RemotePort
        
        'prevent multiple logons
        For i = 2 To lastUser
            If User(i).IPAddress <> "" Then
                
                If User(i).IPAddress = GetIP Then
                    'remove existing name from list
                    For j = 0 To lstUsers.ListCount - 1
                        If User(i).UserName = lstUsers.List(j) Then
                            lstUsers.RemoveItem j
                            Exit For
                        End If
                    Next
                    
                    User(i).IPAddress = ""
                    Exit For
                Else
                    If User(i).UserName = Split(data, Chr(2))(0) Then
                        Winsock1.SendData "ERR:You are already logged on another computer.  If you experienced" & vbCrLf & "a connection dropout, please wait 60 seconds, then try again."
                        Exit Sub
                    End If
                End If
            
            End If
        Next
        
        Set UserRS = QueryDB("SELECT * FROM RegUsers WHERE UserName='" & DataArray(0) & "'")
        
        If UserRS.RecordCount = 0 Then
            Winsock1.SendData "ERR:The username """ & DataArray(0) & """ was not found."
            Exit Sub
        End If
        
        If UserRS("Password") <> DataArray(1) Then
            Winsock1.SendData "ERR:The password you typed was incorrect."
            Exit Sub
        End If
        
        'find the first unused userslot
        n = 2
        While User(n).IPAddress <> ""
            'goto next user record
            n = n + 1
            If n = maxUsers Then
                Winsock1.SendData "ERR:The server is full, please try again later."
                Exit Sub
            End If
        Wend
        
        If n > lastUser Then lastUser = n
        
        'get users registered name
        GetName = UserRS("UserName")
        
        'fill in user's record
        With User(n)
            .UserName = UserRS("UserName")
            .UserID = UserRS("UserID")
            .Timeout = 120
            .IPAddress = GetIP
            .Port = GetPort
            Winsock1.SendData "STA:ONL:" & n
            'add name to listbox
            lstUsers.AddItem .UserName
        End With
        
        
        AddStatus "*** " & User(n).UserName & " has logged on"
        
        UserRS.Close
        
        'notify all connected users that user has logged on
        For i = 2 To lastUser
            If User(i).IPAddress <> "" And i <> n Then
                SendtoUser User(i), "USR:JOI:" & User(n).UserName
            End If
        Next
        
    'Commands (during connection)
    Case "CMD"
        RemoveHeader data
        
        Select Case Left(data, 3)
        Case "FRN"
            'Request for friend list
            RemoveHeader data
            
            RequesterID = User(data).UserID
            
            Winsock1.SendData "STA:FRN:" & MakeFriendList(RequesterID)
            
        Case "LST"
            'request user list
            WhoIsOnline = ""
            For i = 0 To lastUser
                'send user his own name (for testing only)
                If User(i).IPAddress <> "" Then
                'dont send user his own name
                'If User(i).IPAddress <> "" And i <> Val(Split(data, Chr(2))(1)) Then
                    WhoIsOnline = WhoIsOnline & User(i).UserName & Chr(2)
                End If
            Next
            WhoIsOnline = Left(WhoIsOnline, Len(WhoIsOnline) - 1)
            
            Winsock1.SendData "STA:LST:" & WhoIsOnline
        
        'request OffLine Message
        Case "OLM"
            RemoveHeader data
            
            DataArray = Split(data, Chr(2))
            
            Requester = User(DataArray(0)).UserName
            'get user's registered ID
            RequesterID = User(DataArray(0)).UserID
            
            Select Case DataArray(1)
            Case "NEW"
                Set UserRS = QueryDB("SELECT * From OfflineMsgs WHERE Recipient=" & RequesterID & " AND [Read] = FALSE ORDER BY Timestamp DESC")
            Case "ALL"
                Set UserRS = QueryDB("SELECT * From OfflineMsgs WHERE Recipient=" & RequesterID & " ORDER BY Timestamp DESC")
            End Select
            
            'send offline messages
            If UserRS.RecordCount > 0 Then
                UserRS.MoveFirst
                While Not UserRS.EOF
                    
                    Set tempRS = QueryDB("SELECT * FROM RegUsers WHERE UserID=" & UserRS("Sender") & "")
                    If tempRS.RecordCount > 0 Then
                        SendtoUser User(DataArray(0)), "STA:OLM:" & tempRS("UserName") & Chr(2) & UserRS("Timestamp") & Chr(2) & UserRS("Message") & Chr(2) & UserRS("MsgID") & Chr(2) & IIf(UserRS("Read"), "T", "F")
                        tempRS.Close
                    End If
                    
                    UserRS.MoveNext
                Wend
                UserRS.Close
            End If
                
        'FLAG offline message as "Read"
        Case "FLG"
            RemoveHeader data
            
            Set UserRS = QueryDB("SELECT * From OfflineMsgs WHERE MsgID=" & data & "")
            If UserRS.RecordCount > 0 Then
                UserRS("Read") = True
                UserRS.Update
                UserRS.Close
            End If
        
        End Select
    
    Case "DIS"
        'DISconnect user
        RemoveHeader data
        
        Dim userleft As Integer
        userleft = data
        
        For i = 2 To lastUser
            If User(i).IPAddress <> "" Then
                SendtoUser User(i), "USR:PRT:" & User(userleft).UserName
            End If
        Next
        
        For i = 0 To lstUsers.ListCount - 1
            If User(userleft).UserName = lstUsers.List(i) Then
                lstUsers.RemoveItem i
                Exit For
            End If
        Next
        
        AddStatus "*** " & User(userleft).UserName & " has logged off"
        
        With User(userleft)
            .IPAddress = ""
            .Port = ""
            .UserName = ""
        End With
        
        If userleft = lastUser Then lastUser = lastUser - 1
        
    Case "BUZ"
        'MeSsaGe
        Dim dest As String
        Dim src As String
        Dim msg  As String
        RemoveHeader data
            
        src = Split(data, Chr(2))(0)
        dest = Split(data, Chr(2))(1)
        
        'find destination user
        
        'special case for server
        If dest = "@Server" Then
        End If
        
        'look for user with corresponding name
        For i = 2 To lastUser
            If User(i).UserName = dest Then
                SendtoUser User(i), "USR:BUZ:" & src
                Exit For
             End If
        Next
    
    Case "MSG"
        'MeSsaGe
        RemoveHeader data
        
        DataArray = Split(data, Chr(2))
        
        'find destination user
        
        'special case for server
        If DataArray(1) = "@Server" Then
            
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = DataArray(0) Then
                    ChatWindow(i).AddText DataArray(0) & "> " & DataArray(2), CLng(DataArray(3)), DataArray(4), CInt(DataArray(5))
                    bFound = True
                End If
            Next
            
            If Not bFound Then
                Set newChat = New frmChat
                newChat.Caption = DataArray(0)
                newChat.Show
                ChatWindow.Add newChat
                newChat.AddText DataArray(0) & "> " & DataArray(2), CLng(DataArray(3)), DataArray(4), CInt(DataArray(5))
                'newChat.mnuSendFile.Enabled = False
            End If

            Exit Sub
        End If
        
        If Left(DataArray(1), 1) = "@" Then
            
            For i = 1 To ChatWindow.Count
                If ChatWindow(i).Caption = DataArray(1) Then
                    ChatWindow(i).AddText DataArray(0) & "> " & DataArray(2), CLng(DataArray(3)), DataArray(4), CInt(DataArray(5))
                    bFound = True
                End If
            Next
            
            If Not bFound Then
                Set newChat = New frmChat
                newChat.Caption = DataArray(1)
                newChat.Show
                ChatWindow.Add newChat
                newChat.AddText DataArray(0) & "> " & DataArray(2), CLng(DataArray(3)), DataArray(4), CInt(DataArray(5))
                'newChat.mnuSendFile.Enabled = False
                newChat.mnuBuzz.Enabled = False
            End If
        
        End If
        
        
        If DataArray(1) = "@Everyone" Then
            
            'Send to all
            For i = 2 To lastUser
                SendtoUser User(i), "USR:MSG:@Everyone" & Chr(2) & DataArray(0) & "> " & DataArray(2) & Chr(2) & DataArray(3) & Chr(2) & DataArray(4) & Chr(2) & DataArray(5)
            Next

            Exit Sub
        End If
        
        'look for user with corresponding name
        For i = 2 To lastUser
            If User(i).UserName = DataArray(1) Then
                SendtoUser User(i), "USR:MSG:" & DataArray(0) & Chr(2) & DataArray(2) & Chr(2) & DataArray(3) & Chr(2) & DataArray(4) & Chr(2) & DataArray(5)
                Exit For
             End If
        Next
                
    Case "PIC"
        'PICture command
        RemoveHeader data
        
        Select Case Left(data, 3)
        Case "CRC"
            RemoveHeader data
            
            DataArray = Split(data, Chr(2))
            
            RequesterID = DataArray(1)
            
            If GetCRC(App.Path & "\Pictures\" & DataArray(2) & ".jpg") <> DataArray(0) Then
                SendPicture RequesterID, DataArray(2), 0
            End If
            
        Case "REQ"
            'REQuest picture
            RemoveHeader data
            
            DataArray = Split(data, Chr(2))
            
            RequesterID = DataArray(0)
            
            SendPicture RequesterID, DataArray(1), 0
            
        Case "GOT"
            'acknowledge picture data and request next packet
            RemoveHeader data
            
            DataArray = Split(data, Chr(2))
            RequesterID = DataArray(1)
            
            SendPicture RequesterID, DataArray(2), CLng(DataArray(0)) + 1
        
        Case "UPP"
            RemoveHeader data

            Dim tempname As String
            Dim namelen As Long
            Dim partno As Long
            Dim partof As Long
            Dim fname As String
            
            partno = Val(Left(data, 3))
            data = Right(data, Len(data) - 3)
            
            partof = Val(Left(data, 3))
            data = Right(data, Len(data) - 3)
            
            namelen = Val(Left(data, 3))
            data = Right(data, Len(data) - 3)
            
            tempname = Left(data, namelen)
            
            data = Right(data, Len(data) - namelen - 1)
            
            
            fname = Right(tempname, Len(tempname) - InStrRev(tempname, "\"))
            
            Open App.Path & "\Pictures\" & fname For Binary As 1
                Put #1, (partno * 7168) + 1, data
            Close
            
            'Sleep 10
            If partno < partof Then
                Sleep 10
                Winsock1.SendData "PIC:GOT:" & partno & Chr(2) & tempname
            Else
                AddStatus Left(fname, InStr(1, fname, ".") - 1) & " has uploaded a picture."
                Winsock1.SendData "NTF:Your picture was successfully uploaded."
            End If
            
        End Select
        
    End Select
End Sub

Public Sub AddStatus(Text As String)
    frmStatus.AddStatus Text
End Sub

Private Sub SendPicture(RequesterID As Long, picname As String, part As Long)
Dim NumParts As Long
Dim PartData As String
Dim picData As String

    If Dir(App.Path & "\Pictures\" & picname & ".jpg") <> "" Then
        
        Open App.Path & "\Pictures\" & picname & ".jpg" For Binary As 1
            picData = String(LOF(1), Chr(0))
            Get #1, , picData
        Close
        
        NumParts = Len(picData) \ 7168
        
        If Len(picData) Mod 7168 > 0 Then
            NumParts = NumParts + 1
        End If
        
        PartData = String(7168, Chr(0))
        
        PartData = Mid(picData, part * 7168 + 1, 7168)
        SendtoUser User(RequesterID), "PIC:SND:" & Format(part, "000") & Format(NumParts, "000") & Format(Len(picname), "000") & picname & Chr(2) & PartData
        
        picData = ""
    End If
End Sub

Private Function MakeFriendList(RequesterID As Long) As String
Dim FriendList As String
    FriendList = ""
    
    'Build users friends list from database
    Set UserRS = QueryDB("SELECT [UserName] From RegUsers WHERE UserID = ANY (SELECT FriendID FROM FriendsList WHERE UserID=" & RequesterID & ")")
    If UserRS.RecordCount > 0 Then
        UserRS.MoveFirst
        While Not UserRS.EOF
            FriendList = FriendList & UserRS("UserName") & Chr(2)
            UserRS.MoveNext
        Wend
        UserRS.Close
        
        MakeFriendList = Left(FriendList, Len(FriendList) - 1)
    Else
        MakeFriendList = ""
    End If
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


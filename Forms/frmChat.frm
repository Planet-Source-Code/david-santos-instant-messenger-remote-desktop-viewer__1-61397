VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmChat 
   Caption         =   "Chat Window"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9165
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBuzzing 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8160
      Top             =   4800
   End
   Begin VB.Timer tmrBuzz 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   8640
      Top             =   4800
   End
   Begin RichTextLib.RichTextBox rtxChat 
      Height          =   5295
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9340
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":0CCA
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pgbDownload 
      Height          =   135
      Left            =   7320
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5715
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   7320
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   3465
      IntegralHeight  =   0   'False
      Left            =   7320
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   5400
      Width           =   9135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSaveBuffer 
         Caption         =   "Sa&ve Buffer"
      End
      Begin VB.Menu mnuClearBuffer 
         Caption         =   "Clear &Buffer"
      End
      Begin VB.Menu mnuBlank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendFile 
         Caption         =   "&Send a File"
      End
      Begin VB.Menu mnuBuzz 
         Caption         =   "&Buzz!"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuBlank1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font..."
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color..."
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Stores up to 30 of the the user's previous messages
Dim History(30) As String
'our current location in history
Dim curmsg As Integer

'used to check if buzzing is allowed
Property Get Buzzing() As Boolean
    Buzzing = tmrBuzz.Enabled
End Property

Public Sub AddText(data As String, Optional Color As Long = 0, Optional mFont As String = "Arial", Optional size As Integer = 12)
    'display a formatted message in the window
    rtxChat.SelStart = Len(rtxChat.Text)
    rtxChat.SelColor = Color
    rtxChat.SelFontName = mFont
    rtxChat.SelFontSize = size
    rtxChat.SelText = rtxChat.SelText & data & vbCrLf
    rtxChat.SelStart = Len(rtxChat.Text)
End Sub

Private Sub Form_Activate()
    
    If Picture1.Picture = 0 Then
        Dim userPic As String
        userPic = BufferFolder & "\" & Me.Caption & ".jpg"
        'do we have a picture of this user yet?
        If Dir(userPic) = "" Then
            'no, so request a download from the server
            frmMain.Winsock1.SendData "PIC:REQ:" & UserIndex & Chr(2) & Me.Caption
        Else
            'yes, get the CRC of the picture and compare with the server's version
            frmMain.Winsock1.SendData "PIC:CRC:" & GetCRC(userPic) & Chr(2) & UserIndex & Chr(2) & Me.Caption
            Picture1.Picture = LoadPicture(userPic)
        End If
    End If
End Sub

Private Sub Form_Load()
    'get the user's font settings
    CommonDialog1.FontName = mFontName
    CommonDialog1.FontSize = mFontSize
    CommonDialog1.Color = mColor
    
    curmsg = 30
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    'limit size
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 5000 Then Me.Height = 5000
    'adjust controls
    Picture1.Left = Me.Width - Picture1.Width - 8 * 15
    rtxChat.Width = Me.Width - Picture1.Width - 10 * 15
    rtxChat.Height = Me.Height - txtSend.Height - StatusBar1.Height - 56 * 15
    txtSend.Top = rtxChat.Top + rtxChat.Height + 2 * 15
    txtSend.Width = Me.Width - 10 * 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'remove this window from the chatwindows collection
    frmMain.RemoveChatWindow Me.Caption
End Sub

Private Sub mnuBuzz_Click()
    'send buzz command to user
    
    If Not tmrBuzzing.Enabled Then
        frmMain.Winsock1.SendData "BUZ:" & Me.Tag & Chr(2) & Me.Caption
        AddText "***BUZZ!", RGB(255, 0, 0), , 12
        tmrBuzzing.Enabled = True
    End If
End Sub

Public Sub Buzz()
Dim saveX As Long, saveY As Long
Dim i As Integer

    'make sure we can only buzz this window once every 10 seconds
        
    'buzz will reset when timer fires
    tmrBuzz.Enabled = True
    
    'save the current position of this window
    saveY = Me.Top
    saveX = Me.Left

    'do the jiggly!
    For i = 0 To 50
        Me.Top = Me.Top + (Int(Rnd() * 20) - 10) * 20
        Me.Left = Me.Left + (Int(Rnd() * 20) - 10) * 20
        'since this is a loop, we need to prevent it
        'from taking all the processing time
        'so DoEvents gives processing time to do other things
        DoEvents
        'put it back in it's rightful place
        Me.Top = saveY
        Me.Left = saveX
        DoEvents
    Next
        
End Sub

Private Sub mnuClearBuffer_Click()
    'cleat the chat area
    If MsgBox("Clear chat session?", vbQuestion + vbYesNo + _
     vbDefaultButton2, "Chat Session") = vbYes Then rtxChat.Text = ""
End Sub

Private Sub mnuColor_Click()
    'choose a color for your text and save it
    CommonDialog1.ShowColor
    mColor = CommonDialog1.Color
    WriteINI App.path & "\config.ini", "Chat", "Color", CStr(mColor)
End Sub

Private Sub mnuFont_Click()
    'choose a font for your text and save it
    ' .Flags needs to be set to something
    ' before .ShowFont can be used (see MSDN)
    CommonDialog1.Flags = cdlCFScreenFonts
    CommonDialog1.ShowFont
    mFontName = CommonDialog1.FontName
    mFontSize = CommonDialog1.FontSize
    WriteINI App.path & "\config.ini", "Chat", "Font", mFontName
    WriteINI App.path & "\config.ini", "Chat", "FontSize", CStr(mFontSize)

End Sub

Private Sub mnuSaveBuffer_Click()
    If MsgBox("Would you like to save the current chat session?", vbQuestion + vbYesNo, "Chat Session") = vbYes Then
        MakeSureDirectoryExists App.path & "\Logs"
        'save the log as "<name> (<mm-dd-yy>_<hhmm>AM/PM).txt" e.g.: Foobar (7-7-04 0945AM).txt
        rtxChat.SaveFile LogsFolder & "\" & Me.Caption & " (" & Format(Date, "m-d-yy") & "_" & Format(Time, "hhmmAMPM") & ").txt", rtfText
        MsgBox "Chat session saved in ""Logs"" folder", vbInformation, "Chat Session"
    End If
End Sub

Private Sub mnuSendFile_Click()
Dim fSend As New frmSendFile
    'the reason why we create a new frmSend is so that
    ' we can open lots of frmSend windows at the same time
    fSend.Tag = Me.Caption
    fSend.Label1.Caption = "Select a file to send to " & Me.Caption
    fSend.Show , Me
End Sub

Private Sub tmrBuzz_Timer()
    'this will allow the user to be buzzed again
    tmrBuzz.Enabled = False
End Sub

Private Sub tmrBuzzing_Timer()
    tmrBuzzing.Enabled = False
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

    If KeyCode = vbKeyReturn Then
        If Trim(txtSend.Text) = "" Then Exit Sub
        
        'echo user's message to screen
        If (Me.Caption = "@Server") Or (Left(Me.Caption, 1) <> "@") Then
            'Add text if talking to server or another user.
            'if we add text while talking to everybody, we will get two
            'copies of our message since the server sends the message to everyone
            'including the sender
            AddText UserName & "> " & txtSend.Text, mColor, mFontName, mFontSize
        End If
        
        'and send it to the server for forwarding
        frmMain.Winsock1.SendData "MSG:" & Me.Tag & Chr(2) & Me.Caption & Chr(2) & txtSend.Text & Chr(2) & mColor & Chr(2) & mFontName & Chr(2) & mFontSize
        
        'save the message in history if its new
        If curmsg = 30 Then History(30) = txtSend.Text
        
        'copy the History buffer upwards
        'after we type 30 times, we lose whatever we first typed
        For i = 0 To 29
            History(i) = History(i + 1)
        Next
        
        'clear the text so we can type again
        txtSend.Text = ""
        txtSend.SetFocus
        
        curmsg = 30
    End If
    
    If KeyCode = vbKeyUp Then
        'go back once in history
        curmsg = curmsg - 1
        If curmsg < 0 Then curmsg = 0
        If History(curmsg) = "" Then curmsg = curmsg + 1
        'display it
        txtSend.Text = History(curmsg)
        txtSend.SelStart = 0
        txtSend.SelLength = Len(txtSend.Text)
        'prevent the up key from having an effect in the textbox
        KeyCode = 0
    End If

    If KeyCode = vbKeyDown Then
        'go forward once in history
        curmsg = curmsg + 1
        If curmsg > 30 Then curmsg = 30
        'display it
        txtSend.Text = History(curmsg)
        txtSend.SelStart = 0
        txtSend.SelLength = Len(txtSend.Text)
        'prevent the down key from having an effect in the textbox
        KeyCode = 0
    End If
End Sub

Public Sub SendFile()
    mnuSendFile_Click
End Sub

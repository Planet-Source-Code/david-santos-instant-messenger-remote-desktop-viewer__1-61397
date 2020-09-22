VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmChat 
   Caption         =   "Chat Window"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9165
   Icon            =   "frmChatServ.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   9165
   Begin VB.Timer tmrBuzz 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7800
      Top             =   4800
   End
   Begin RichTextLib.RichTextBox rtxChat 
      Height          =   5295
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9340
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmChatServ.frx":0CCA
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pgbDownload 
      Height          =   135
      Left            =   7320
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   0
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
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MsgBuffer(30) As String
Dim CurMsg As Integer
Dim bBuzz As Boolean

Property Get Buzzing() As Boolean
    Buzzing = bBuzz
End Property

Public Sub AddText(data As String, Optional Color As Long = 0, Optional mFont As String = "Arial", Optional size As Integer = 12)
    rtxChat.SelStart = Len(rtxChat.Text)
    rtxChat.SelColor = Color
    rtxChat.SelFontName = mFont
    rtxChat.SelFontSize = size
    rtxChat.SelText = rtxChat.SelText & data & vbCrLf
    rtxChat.SelStart = Len(rtxChat.Text)
End Sub

Private Sub Form_Activate()
Dim userPic As String
    
    If Picture1.Picture = 0 Then
        ' checking if .picture = 0 is a way to check if a picture has been loaded
        ' the reason we do this is because every time we click on the form,
        ' the Form_Activate event fires. We don't want to reload the picture
        ' everytime we switch to the form, so we make sure it only happens once
        
        ' The reason why we didn't put picture loading into the Form_Load event
        ' is because the Me.Caption property has not been set during load
        
        userPic = App.Path & "\Pictures\" & Me.Caption & ".jpg"
        If Dir(userPic) <> "" Then
            Picture1.Picture = LoadPicture(userPic)
        End If
    End If

End Sub

Private Sub Form_Load()
    CommonDialog1.FontName = mFontName
    CommonDialog1.FontSize = mFontSize
    CommonDialog1.Color = mColor
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 5000 Then Me.Width = 5000
    If Me.Height < 5000 Then Me.Height = 5000
    Picture1.Left = Me.Width - Picture1.Width - 10 * 15
    rtxChat.Width = Me.Width - Picture1.Width - 12 * 15
    rtxChat.Height = Me.Height - txtSend.Height - StatusBar1.Height - 40 * 15
    txtSend.Top = rtxChat.Top + rtxChat.Height + 2 * 15
    txtSend.Width = Me.Width - 10 * 15
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mdiServer.RemoveChatWindow Me.Caption
End Sub



Private Sub mnuTileH_Click()
    mdiServer.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileV_Click()
    mdiServer.Arrange vbTileVertical
End Sub


Private Sub mnuCascade_Click()
    mdiServer.Arrange vbCascade
End Sub


Private Sub mnuArrange_Click()
    mdiServer.Arrange vbArrangeIcons
End Sub


Private Sub mnuBuzz_Click()
    mdiServer.SendtoName Me.Caption, "USR:BUZ:@Server"
    AddText "***BUZZ!", RGB(255, 0, 0), , 12
End Sub


'Question: can the server be buzzed?
Public Sub Buzz()
Dim saveX As Long, saveY As Long
Dim i As Integer

    bBuzz = True
    saveY = Me.Top
    saveX = Me.Left

    For i = 0 To 50
        Me.Top = Me.Top + (Int(Rnd() * 20) - 10) * 20
        Me.Left = Me.Left + (Int(Rnd() * 20) - 10) * 20
        DoEvents
        Me.Top = saveY
        Me.Left = saveX
        DoEvents
    Next
    tmrBuzz.Enabled = True
End Sub

Private Sub mnuClearBuffer_Click()
    rtxChat.Text = ""
End Sub

Private Sub mnuColor_Click()
    CommonDialog1.ShowColor
    mColor = CommonDialog1.Color
    WriteINI App.Path & "\config.ini", "Chat", "Color", CStr(mColor)
End Sub

Private Sub mnuFont_Click()
    CommonDialog1.Flags = cdlCFScreenFonts
    CommonDialog1.ShowFont
    mFontName = CommonDialog1.FontName
    mFontSize = CommonDialog1.FontSize
    WriteINI App.Path & "\config.ini", "Chat", "Font", mFontName
    WriteINI App.Path & "\config.ini", "Chat", "FontSize", CStr(mFontSize)
End Sub

Private Sub tmrBuzz_Timer()
    bBuzz = False
    tmrBuzz.Enabled = False
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
    If KeyCode = vbKeyReturn Then
        If Trim(txtSend.Text) = "" Then Exit Sub
            
        'echo user's message to screen
        
        AddText "@Server> " & txtSend.Text, mColor, mFontName, mFontSize
        
        If Left(Me.Caption, 1) <> "@" Then
            mdiServer.SendtoName Me.Caption, "USR:MSG:@Server" & Chr(2) & txtSend.Text & Chr(2) & mColor & Chr(2) & mFontName & Chr(2) & mFontSize
        Else
            mdiServer.SendtoName Me.Caption, "USR:MSG:@Everyone" & Chr(2) & "@Server> " & txtSend.Text & Chr(2) & mColor & Chr(2) & mFontName & Chr(2) & mFontSize
        End If
        
        If CurMsg = 30 Then MsgBuffer(30) = txtSend.Text
        
        For i = 0 To 29
            MsgBuffer(i) = MsgBuffer(i + 1)
        Next
        
        txtSend.Text = ""
        txtSend.SetFocus
        
        CurMsg = 30
    
    End If
    
    If KeyCode = vbKeyUp Then
        CurMsg = CurMsg - 1
        If CurMsg < 0 Then CurMsg = 0
        If MsgBuffer(CurMsg) = "" Then CurMsg = CurMsg + 1
        txtSend.Text = MsgBuffer(CurMsg)
        txtSend.SelStart = 0
        txtSend.SelLength = 0
        KeyCode = 0
    End If

    If KeyCode = vbKeyDown Then
        CurMsg = CurMsg + 1
        If CurMsg > 30 Then CurMsg = 30
        txtSend.Text = MsgBuffer(CurMsg)
        txtSend.SelStart = 0
        txtSend.SelLength = 0
        KeyCode = 0
    End If
End Sub




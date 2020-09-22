VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRegistration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration Wizard"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   ControlBox      =   0   'False
   Icon            =   "frmRegistration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4920
      Left            =   120
      Picture         =   "frmRegistration.frx":0CCA
      ScaleHeight     =   4920
      ScaleWidth      =   2190
      TabIndex        =   33
      Top             =   120
      Width           =   2190
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "&Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   5160
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picStep 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   2
      Left            =   2400
      ScaleHeight     =   4920
      ScaleWidth      =   7215
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtID 
         Height          =   285
         Left            =   2760
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2760
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter your account information here and click ""Next"" to proceed with registration"
         Height          =   195
         Left            =   720
         TabIndex        =   29
         Top             =   960
         Width           =   5685
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRegistration.frx":240CC
         Height          =   495
         Left            =   1080
         TabIndex        =   30
         Top             =   3120
         Width           =   5055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Student Number"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Password*"
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
   End
   Begin VB.PictureBox picStep 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   3
      Left            =   2400
      ScaleHeight     =   4920
      ScaleWidth      =   7215
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      Begin VB.PictureBox picPass 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   120
         ScaleHeight     =   975
         ScaleWidth      =   4095
         TabIndex        =   37
         Top             =   3120
         Visible         =   0   'False
         Width           =   4095
         Begin VB.TextBox txtVerifyPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1440
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   22
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtNewPass 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1440
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&Verify Password"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "&New Password"
            Height          =   255
            Left            =   0
            TabIndex        =   19
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   4920
         TabIndex        =   17
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   12
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txteMail 
         Height          =   285
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox txtOther 
         Height          =   765
         Left            =   1560
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1560
         Width           =   2895
      End
      Begin VB.PictureBox picMyPic 
         Height          =   1815
         Left            =   4920
         ScaleHeight     =   1755
         ScaleWidth      =   1755
         TabIndex        =   31
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   1560
         TabIndex        =   23
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CheckBox chkChangePass 
         Caption         =   "Change &password"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Use the Browse button to upload a picture that you would like others will see when they chat with you."
         Height          =   855
         Left            =   4920
         TabIndex        =   38
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Name"
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&E-mail"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Other"
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "This is your personal information that others will be able to view. To update your information click the Update button below"
         Height          =   495
         Left            =   960
         TabIndex        =   32
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.PictureBox picStep 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   0
      Left            =   2400
      ScaleHeight     =   4920
      ScaleWidth      =   7215
      TabIndex        =   24
      Top             =   120
      Width           =   7215
      Begin VB.OptionButton Option2 
         Caption         =   "My account is already active"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   2880
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "I do not have a login name yet"
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please select an option and click Next to proceed."
         Height          =   195
         Left            =   3240
         TabIndex        =   41
         Top             =   3720
         Width           =   3570
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "If this is your first time you may choose a login name for your account."
         Height          =   315
         Left            =   480
         TabIndex        =   36
         Top             =   1680
         Width           =   5730
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRegistration.frx":2415D
         Height          =   435
         Left            =   480
         TabIndex        =   35
         Top             =   1080
         Width           =   5730
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to the Yippee! Messenger Registration Wizard."
         Height          =   195
         Left            =   480
         TabIndex        =   34
         Top             =   600
         Width           =   4005
      End
   End
   Begin VB.PictureBox picStep 
      BorderStyle     =   0  'None
      Height          =   4920
      Index           =   1
      Left            =   2400
      ScaleHeight     =   4920
      ScaleWidth      =   7215
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   6
         Tag             =   "This name will be used on the user lists"
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Click Next to proceed."
         Height          =   195
         Left            =   4560
         TabIndex        =   40
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "(Maximum of 30 chars)"
         Height          =   195
         Left            =   2640
         TabIndex        =   39
         Top             =   2520
         Width           =   1590
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Login Name"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRegistration.frx":241E8
         Height          =   675
         Left            =   720
         TabIndex        =   28
         Top             =   480
         Width           =   4365
      End
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curFrame As Long

Private Sub chkChangePass_Click()
    picPass.Visible = chkChangePass.Value = 1
End Sub

Private Sub cmdBack_Click()
    Select Case curFrame
    Case 3
        picMyPic.Picture = LoadPicture("")
        curFrame = curFrame - 1
    Case 2
        If Option1.Value = True Then curFrame = curFrame - 1
        If Option2.Value = True Then curFrame = curFrame - 2
    Case Else
        curFrame = curFrame - 1
    End Select
    
    If curFrame = 0 Then cmdBack.Enabled = False
    OpenFrame
    cmdNext.Enabled = True
End Sub

'this is called from frmMain when a REG packet is recieved
Public Sub Connected()
    txtName.Text = ""
    txteMail.Text = ""
    txtOther.Text = ""
    txtNewPass = ""
    txtVerifyPass = ""
    cmdBack.Enabled = False
    Command3.Caption = "&Finish"
    curFrame = 3
    cmdNext.Enabled = False
    OpenFrame
End Sub

Private Sub cmdNext_Click()
    Select Case curFrame
    Case 0
        If Option1.Value = True Then
            curFrame = curFrame + 1
        ElseIf Option2.Value = True Then
            curFrame = curFrame + 2
        End If
        OpenFrame
    Case 1
        If Trim(txtLogin.Text) = "" Then
            MsgBox "Please enter a Login name", vbInformation, "No Login name entered"
            Exit Sub
        End If
        curFrame = curFrame + 1
        OpenFrame
    Case 2
        If Trim(txtID.Text) = "" Or Trim(txtPassword.Text) = "" Then
            MsgBox "Please fill in all the required information.", vbInformation, "Error"
            Exit Sub
        End If
        
        frmMain.Winsock1.Close
        frmMain.Winsock1.Bind frmMain.Winsock1.LocalPort, frmMain.Winsock1.LocalIP
        frmMain.Winsock1.SendData "REG:" & Trim(txtLogin.Text) & Chr(2) & Trim(txtID.Text) & Chr(2) & Trim(LCase(txtPassword.Text))

    End Select
    
    cmdBack.Enabled = True
End Sub

Private Sub OpenFrame()
Dim i As Integer
    For i = 0 To 3
        picStep(i).Visible = False
    Next
    picStep(curFrame).Visible = True
    Select Case curFrame
    Case 1
        txtLogin.SetFocus
    Case 2
        txtID.SetFocus
    End Select
End Sub

Private Sub cmdBrowse_Click()
On Local Error GoTo errUpdatePic
Dim fname As String
    
    ' set this to true to generate an error when we click cancel thile browsing
    
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "JPEG image files|*.jpg"
    CommonDialog1.ShowOpen
    
    ' if we don't set the cancelerror to True, then the
    ' program will continue here even if we press cancel...
    
    If MsgBox("Do you want to update your picture at the server with this file?", vbYesNo + vbQuestion, "Update Picture") = vbYes Then
        fname = BufferFolder & "\" & UserName & ".jpg"
        
        If FileLen(CommonDialog1.FileName) > 204800 Then
            MsgBox "Please limit you picture files to 200KB.", vbInformation, "File too large"
            Exit Sub
        End If
        
        FileCopy CommonDialog1.FileName, fname
        picMyPic.Picture = LoadPicture(fname)
        frmMain.SendPic fname, 0
    End If
    
    Exit Sub
    
errUpdatePic:
    ' cancel error will jump here, but since the error
    ' isn't critical, ignore it
End Sub

Private Sub cmdUpdate_Click()
Dim newpass As String

    If Trim(txtName.Text) = "" And Trim(txteMail.Text) = "" And Trim(txtOther.Text) = "" Then
        MsgBox "Please fill in at least one of the fields.", vbInformation, "Registration Error"
        Exit Sub
    End If

    If chkChangePass.Value = 1 Then
        'update the password if change password is checked
        
        If Len(Trim(txtNewPass.Text)) < 8 Then
            MsgBox "Your password must be at least 8 characters.", vbInformation, "Error"
            Exit Sub
        End If
        
        If Trim(LCase(txtNewPass.Text)) = Trim(LCase(txtVerifyPass.Text)) Then
            newpass = Trim(LCase(txtNewPass.Text))
        Else
            MsgBox "The new passwords do not match. Please correct them and try again.", vbInformation, "Error"
            Exit Sub
        End If
        
    End If
    frmMain.Winsock1.SendData "UPD:" & Trim(txtLogin.Text) & Chr(2) & Trim(txtID.Text) & Chr(2) & Trim(txtPassword.Text) & Chr(2) & Trim(txtName.Text) & Chr(2) & Trim(txteMail.Text) & Chr(2) & Trim(txtOther.Text) & Chr(2) & newpass
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    curFrame = 0
    Me.Left = Screen.Width - Me.Width - frmMain.Width - 250
    Me.Top = 500
End Sub

Private Sub Option1_Click()
    Label8.Visible = True
    txtLogin.Text = ""
    txtID.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub Option2_Click()
    Label8.Visible = False
    txtLogin.Text = ""
    txtID.Text = ""
    txtPassword.Text = ""
End Sub

Private Sub txtID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtPassword.SetFocus
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack, vbKeyReturn, vbKeyDelete
        ' allow only numbers and backspace in the ID
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtLogin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdNext_Click
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    ' limit the characters that can be used in a username
    Select Case KeyAscii
    Case Asc("!"), Asc("/"), Asc("\"), Asc("|"), Asc("."), Asc(","), Asc("<"), Asc(">"), Asc("%"), Asc("*"), Asc(":"), Asc(";"), Asc("""")
        'prevent characters that are not allowed in filenames, because the system
        'uses the users name for the picture files
        KeyAscii = 0
    Case Else
        ' allow everything else
    End Select
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdNext_Click
End Sub


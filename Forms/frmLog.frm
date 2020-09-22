VERSION 5.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3615
   ControlBox      =   0   'False
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   405
      Left            =   1440
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   405
      Left            =   1200
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Log In"
      Height          =   405
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Already Registered?"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3375
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Username:"
         Height          =   195
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   840
         Width           =   810
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "No Username yet?"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    If Trim(txtUser.Text) = "" Or Trim(txtPass.Text) = "" Then
        MsgBox "Please enter a login name and a password", vbExclamation, "Login Error"
        Exit Sub
    End If
    UserName = Trim(txtUser.Text)
    Password = Trim(LCase(txtPass.Text))
    frmMain.Connect
    Unload Me
End Sub


Private Sub cmdRegister_Click()
    frmRegistration.Show vbModal
End Sub

Private Sub Form_Activate()
    txtPass.Text = Password
    txtUser.Text = UserName
    
    cmdLogin.SetFocus
    If Password = "" Then txtPass.SetFocus
    If UserName = "" Then txtUser.SetFocus
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width - 150
    Me.Top = 500
End Sub

Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser.Text)
End Sub

Private Sub txtPass_GotFocus()
    txtPass.SelStart = 0
    txtPass.SelLength = Len(txtPass.Text)
End Sub

Private Sub txtUser_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdLogin_Click
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdLogin_Click
End Sub


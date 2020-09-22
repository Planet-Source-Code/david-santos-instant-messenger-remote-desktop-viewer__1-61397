VERSION 5.00
Begin VB.Form frmAboutServer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Yippee! IM Server"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   Icon            =   "frmAboutServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2370
      Left            =   0
      Picture         =   "frmAboutServer.frx":0CCA
      ScaleHeight     =   2370
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4320
         Top             =   1800
      End
   End
End
Attribute VB_Name = "frmAboutServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
Static i As Long
    
    If i > 1 Then
        InvertCase i
    End If
    
    i = i + 1
    If i > Len(Me.Caption) Then i = 1
    
    InvertCase i
End Sub

Private Sub InvertCase(pos As Long)
Dim cap As String
    cap = Me.Caption
    Select Case Asc(Mid$(cap, pos, 1))
    Case Asc("a") To Asc("z")
        Mid$(cap, pos, 1) = UCase(Mid$(cap, pos, 1))
    Case Asc("A") To Asc("Z")
        Mid$(cap, pos, 1) = LCase(Mid$(cap, pos, 1))
    End Select
    Me.Caption = cap
End Sub

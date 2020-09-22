VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add a Friend"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.CommandButton cmdAddFriend 
      Caption         =   "Add User to Friends List"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the user's Account/ID number OR login name, and click the button below to add the user to your Friends List"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddFriend_Click()
    frmMain.Winsock1.SendData "ADF:" & UserIndex & Chr(2) & Trim(txtAdd.Text)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub txtAdd_GotFocus()
    txtAdd.SelStart = 0
    txtAdd.SelLength = Len(txtAdd.Text)
End Sub

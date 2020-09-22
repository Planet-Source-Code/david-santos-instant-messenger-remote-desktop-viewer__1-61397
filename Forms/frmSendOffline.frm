VERSION 5.00
Begin VB.Form frmSendOffline 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Offline Message"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMessage 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txtTo 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmSendOffline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    frmMain.Winsock1.SendData "SOL:" & UserIndex & Chr(2) & txtTo.Text & Chr(2) & Trim(txtMessage.Text)
    MsgBox "The user will recieve the message the next time they sign on.", vbInformation, "Notice"
    Unload Me
End Sub

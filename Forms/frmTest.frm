VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "0"
   ClientHeight    =   4500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4200
      Width           =   7575
   End
   Begin ChatClient.ucChatBox ucChatBox1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    ucChatBox1.Width = Me.Width - 90
    ucChatBox1.Height = Me.Height - Text1.Height - 60 - 350
    Text1.Top = ucChatBox1.Height + 30
    Text1.Width = Me.Width - 90
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ucChatBox1.Add Text1.Text
        Text1.Text = ""
    End If
End Sub

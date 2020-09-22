VERSION 5.00
Begin VB.Form frmStatus 
   Caption         =   "Status"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4575
   ScaleWidth      =   5910
   Begin VB.TextBox txtStatus 
      Height          =   4575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub AddStatus(Text As String)
    txtStatus.Text = txtStatus.Text & Text & vbCrLf
    txtStatus.SelStart = Len(txtStatus.Text)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'prevent user from closing the window from the close button
    'window will still close if Windows shuts down, or if the server unloads
    If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 420 Then Exit Sub
    txtStatus.Width = Me.Width - 8 * 15
    txtStatus.Height = Me.Height - 34 * 15
End Sub


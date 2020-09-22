VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOfflineMsg 
   Caption         =   "Offline Messages"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmOfflineMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMessage 
      Height          =   1575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2040
      Width           =   5775
   End
   Begin MSComctlLib.ListView lvwOfflineMsgs 
      Height          =   1935
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3413
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sender"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Message"
         Object.Width           =   6068
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3690
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOfflineMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    lvwOfflineMsgs.Width = Me.Width - 150
    lvwOfflineMsgs.Height = Me.Height - txtMessage.Height - 750
    txtMessage.Top = lvwOfflineMsgs.Height + 50
    txtMessage.Width = Me.Width - 150
End Sub

Private Sub lvwOfflineMsgs_Click()
    With lvwOfflineMsgs.SelectedItem
        txtMessage.Text = "Sender:" & vbTab & .Text & vbCrLf & "Sent:" & vbTab & .SubItems(1) & vbCrLf & vbCrLf & .SubItems(2)
        If .Bold Then
            frmMain.Winsock1.SendData "CMD:FLG" & Chr(2) & .Tag
            .Bold = False
            .ListSubItems(1).Bold = False
            .ListSubItems(2).Bold = False
        End If
    End With
End Sub


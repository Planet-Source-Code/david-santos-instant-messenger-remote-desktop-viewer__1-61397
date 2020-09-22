VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSendFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send a File"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmSendFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescription 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a file to send to %1"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
On Error GoTo errCancel
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    txtFileName.Text = CommonDialog1.FileName
    Exit Sub
errCancel:
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
Dim newPort As Long
    'start with the base transfer port
    newPort = baseFileTransferPort
    
    'create a File Transfer Form in download mode
    If frmMain.NewTransfer(CommonDialog1.FileName, FileLen(CommonDialog1.FileName), Me.Tag, newPort, FTR_UPLOAD) Then
        'Send an FTR to the user via UDP
        'if port was updated during NewTransfer, it will be reflected here
        frmMain.Winsock1.SendData "FIL:REQ:" & UserName & Chr(2) & Me.Tag & Chr(2) & CommonDialog1.FileTitle & Chr(2) & FileLen(CommonDialog1.FileName) & Chr(2) & txtDescription.Text & Chr(2) & frmMain.Winsock1.LocalIP & Chr(2) & newPort
    End If
    
    Unload Me
End Sub

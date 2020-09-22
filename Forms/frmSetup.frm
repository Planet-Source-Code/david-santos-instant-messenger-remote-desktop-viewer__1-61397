VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection settings"
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4695
      Begin VB.CheckBox chkAutorun 
         Caption         =   "&Auto Run"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtLocalPort 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtRemotePort 
         Height          =   285
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incoming Port"
         Height          =   195
         Left            =   1155
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outgoing Port"
         Height          =   195
         Left            =   1155
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name/IP Address"
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   1770
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAutorun_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
        
    If UserIndex > -1 Then
        If MsgBox("You will need to re-connect after changing these settings. Continue?", vbQuestion + vbYesNo, "Accept changes?") = vbNo Then
            Exit Sub
        End If
        frmMain.Disconnect
    End If
    
    frmMain.Winsock1.Close
    
    frmMain.Winsock1.RemoteHost = Trim(txtServerIP.Text)
    frmMain.Winsock1.RemotePort = Trim(txtRemotePort.Text)
    frmMain.Winsock1.LocalPort = Trim(txtLocalPort.Text)

    If chkAutorun.Value = 1 Then
        SetAutoRun True
    Else
        SetAutoRun False
    End If

    WriteINI App.path & "\config.ini", "Settings", "RemoteHost", frmMain.Winsock1.RemoteHost
    WriteINI App.path & "\config.ini", "Settings", "RemotePort", frmMain.Winsock1.RemotePort
    WriteINI App.path & "\config.ini", "Settings", "LocalPort", frmMain.Winsock1.LocalPort

    cmdApply.Enabled = False

End Sub

Private Sub cmdApply_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cmdApply.Enabled Then cmdApply_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width - frmMain.Width - 250
    Me.Top = 500
    txtServerIP.Text = frmMain.Winsock1.RemoteHost
    txtRemotePort.Text = frmMain.Winsock1.RemotePort
    txtLocalPort.Text = frmMain.Winsock1.LocalPort
    chkAutorun.Value = IIf(GetAutoRun, 1, 0)
End Sub


Private Sub txtLocalPort_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub txtRemotePort_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub txtRemotePort_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyDelete, vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtServerIP_KeyDown(KeyCode As Integer, Shift As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub txtServerIP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9, Asc("."), vbKeyDelete, vbKeyBack
    Case Else
        KeyAscii = 0
    End Select
End Sub

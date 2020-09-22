VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yippee Proxy Server"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtLog 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   600
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtRemotePort 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "6350"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "192.168.185.14"
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox txtLocal 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "8000"
      Top             =   120
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6669
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Server Port"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait on Port"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
    txtLog.Text = ""
End Sub

Private Sub cmdReset_Click()
    Winsock1.Close
    Winsock2.Close
    
    Winsock1.Protocol = sckUDPProtocol
    Winsock2.Protocol = sckUDPProtocol
    
    Winsock1.LocalPort = txtLocal.Text
    Winsock1.Bind
    'Winsock1.Listen

    Winsock2.RemoteHost = txtRemoteIP.Text
    Winsock2.RemotePort = txtRemotePort.Text
End Sub

Private Sub Form_Load()
    cmdReset_Click


End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    'Winsock1.Close
    'Winsock1.Accept requestID
    'Winsock2.Connect txtRemoteIP.Text, txtRemotePort.Text
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
On Error GoTo errh:
    Winsock1.GetData sData
    txtLog.Text = txtLog.Text & "IN:" & sData & vbCrLf
    Winsock2.SendData sData
    Debug.Print "WINSOCK1:" & sData
errh:
    txtLog.Text = txtLog.Text & Err.Description & vbCrLf
End Sub

Private Sub Winsock2_Close()
    Winsock1.Close
    Winsock2.Close
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
On Error GoTo errh:
    Winsock2.GetData sData
    txtLog.Text = txtLog.Text & "OUT:" & sData & vbCrLf
    Winsock1.SendData sData
    Debug.Print "WINSOCK2:" & sData
    Exit Sub
errh:
    txtLog.Text = txtLog.Text & Err.Description & vbCrLf
End Sub


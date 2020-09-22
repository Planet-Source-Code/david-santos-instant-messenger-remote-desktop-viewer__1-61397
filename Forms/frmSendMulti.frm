VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmSendMulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MultiShare Files"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frmSendMulti.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6180
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4920
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5520
      Top             =   3960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select &All"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.ListBox lstRecipients 
      Height          =   2295
      IntegralHeight  =   0   'False
      Left            =   2880
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recipients"
      Height          =   2895
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   5655
      Begin VB.CommandButton cmdInvert 
         Caption         =   "&Invert"
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a &file to share:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmSendMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim curUser As Integer
Dim filedata() As Byte

Private Sub cmdBrowse_Click()
On Error GoTo errCancel
    CommonDialog1.CancelError = True
    CommonDialog1.ShowOpen
    txtFileName.Text = CommonDialog1.fileName
    
    cmdSend.Enabled = True
errCancel:
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    For i = 0 To lstRecipients.ListCount - 1
        lstRecipients.Selected(i) = False
    Next
End Sub

Private Sub cmdInvert_Click()
    For i = 0 To lstRecipients.ListCount - 1
        lstRecipients.Selected(i) = Not lstRecipients.Selected(i)
    Next
End Sub

Private Sub cmdSelectAll_Click()
    For i = 0 To lstRecipients.ListCount - 1
        lstRecipients.Selected(i) = True
    Next
End Sub

Private Sub cmdSend_Click()
    If Dir(CommonDialog1.fileName) = "" Then
        MsgBox "File does not exist. Please check the file again.", vbExclamation, "Multisend Error"
        Exit Sub
    End If
    
    If lstRecipients.SelCount = 0 Then
        MsgBox "No recipient(s) to send to!!!", vbExclamation, "Multisend Error"
        Exit Sub
    End If
    
    cmdSend.Enabled = False
    curUser = 0
    
    Open CommonDialog1.fileName For Binary As 1
    ReDim filedata(LOF(1))
    Get #1, , filedata
    Close 1
    
    Winsock1.Close
    On Error GoTo errBind
    Winsock1.Bind Winsock1.LocalPort, Winsock1.LocalIP
    On Error GoTo 0
    Winsock1.Listen

    'find the first selected user
    While Not lstRecipients.Selected(curUser)
        curUser = curUser + 1
    Wend
        
    mdiServer.SendtoName lstRecipients.List(curUser), "USR:FIL:@Server" & Chr(2) & CommonDialog1.FileTitle & Chr(2) & FileLen(CommonDialog1.fileName) & Chr(2) & "A file from the server" & Chr(2) & Winsock1.LocalIP & Chr(2) & Winsock1.LocalPort
    Exit Sub
errBind:
    Winsock1.LocalPort = Winsock1.LocalPort + 1
    Resume
End Sub

Private Sub Form_Load()
    For i = 0 To mdiServer.lstUsers.ListCount - 1
        lstRecipients.AddItem mdiServer.lstUsers.List(i)
    Next
    Winsock1.LocalPort = 9100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdSend.Enabled Then
        If MsgBox("Multisend has not finished yet. Would you like to cancel?", vbYesNo + vbExclamation + vbDefaultButton2, "Multisend still active") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    Winsock1.Close
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim sData As String
    Winsock1.GetData sData
    If Left(sData, 3) = "OK!" Then Winsock1.SendData filedata
End Sub

Private Sub Winsock1_SendComplete()
    Winsock1.Close
    mdiServer.AddStatus "*** Sent file """ & CommonDialog1.FileTitle & """ to " & lstRecipients.List(curUser)
    If curUser < lstRecipients.ListCount - 1 Then
        curUser = curUser + 1
        'find the next selected user
        While Not lstRecipients.Selected(curUser)
            curUser = curUser + 1
            If curUser > lstRecipients.ListCount - 1 Then Exit Sub
        Wend
        
        Winsock1.Close
        On Error GoTo errBind
        Winsock1.Bind Winsock1.LocalPort, Winsock1.LocalIP
        On Error GoTo 0
        Winsock1.Listen
        
        mdiServer.SendtoName lstRecipients.List(curUser), "USR:FIL:@Server" & Chr(2) & CommonDialog1.FileTitle & Chr(2) & FileLen(CommonDialog1.fileName) & Chr(2) & "A distributed file from the server" & Chr(2) & Winsock1.LocalIP & Chr(2) & Winsock1.LocalPort
    Else
        cmdSend.Enabled = True
        Winsock1.Close
        curUser = 0
    End If
    Exit Sub

errBind:
    Winsock1.LocalPort = Winsock1.LocalPort + 1
    Resume
End Sub

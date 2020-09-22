VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Database"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   Icon            =   "frmDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6840
   Begin VB.Frame Frame1 
      Caption         =   "Add/Remove Users"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   6615
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtUserNo 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Student Number"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdClearMsgs 
      Caption         =   "Delete Offline Messages"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   4440
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5318
      View            =   3
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
         Text            =   "Student Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "UserName"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
On Error GoTo errExists
    
    If Trim(txtUserNo.Text) = "" Then Exit Sub
    
    Set UserRS = QueryDB("SELECT * FROM RegUsers")
    UserRS.AddNew
    UserRS("UserNo") = Trim(txtUserNo.Text)
    'set default password to 12345
    UserRS("Password") = "12345"
    UserRS.Update
    UserRS.Close
    RefreshDB
    txtUserNo.Text = ""
    txtUserNo.SetFocus
    Exit Sub

errExists:
    ' an error will occur if the UserNo already exists
    ' because the Index property of the UserNo field in
    ' the RegUsers table is set to Yes (No duplicates)
    MsgBox "That Student number is already in the database. Please enter another number", vbInformation, "Error Adding User"
End Sub

Private Sub cmdClearMsgs_Click()
Dim msgcnt As Integer
    
    If MsgBox("Are you sure you want to clear all read offline messages?", vbQuestion + vbYesNo, "Confirm Action") = vbYes Then
        'delete all messages maked as read in the database
        Set UserRS = QueryDB("SELECT * FROM OfflineMsgs WHERE [Read] = TRUE")
        msgcnt = UserRS.RecordCount
        While Not UserRS.EOF
            UserRS.Delete
            UserRS.Update
            UserRS.MoveNext
        Wend
        MsgBox msgcnt & " messages were deleted.", vbInformation, "Messages deleted"
        UserRS.Close
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    'make sure something has been selected in the list
    If Not (ListView1.SelectedItem Is Nothing) Then
        If MsgBox("Are you sure you want to remove the user """ & ListView1.SelectedItem.SubItems(1) & """?", vbQuestion + vbYesNo, "Confirm Action") = vbYes Then
            'delete the user from the database
            Set UserRS = QueryDB("SELECT * FROM RegUsers WHERE UserNo = '" & ListView1.SelectedItem.Text & "'")
            UserRS.Delete
            UserRS.Update
            UserRS.Close
            RefreshDB
        End If
    End If
End Sub

Private Sub Form_Load()
    RefreshDB
End Sub

Private Sub RefreshDB()
    'update the list of online users
    ListView1.ListItems.Clear
    Set UserRS = QueryDB("SELECT * FROM RegUsers")
    
    While Not UserRS.EOF
        ListView1.ListItems.Add , , UserRS("UserNo")
        
        ' ignore any errors that will occur when the UserName is Null
        ' This will occur only when the user has not registered yet
        On Error Resume Next
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = UserRS("UserName")
        'resume error checking
        On Error GoTo 0
        
        ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = UserRS("Password")
        UserRS.MoveNext
    Wend

    UserRS.Close
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    'respond to the delete key in the listview
    If KeyAscii = vbKeyDelete Then cmdRemove_Click
End Sub

Private Sub txtUserNo_KeyDown(KeyCode As Integer, Shift As Integer)
    'respond to the Enter key when entering the user number
    If KeyCode = vbKeyReturn Then cmdAdd_Click
End Sub

Private Sub txtUserNo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack, vbKeyDelete, vbKeyReturn
        'asslow these keys to be pressed
    Case Else
        'block everything else by setting the key value to 0
        KeyAscii = 0
    End Select
End Sub

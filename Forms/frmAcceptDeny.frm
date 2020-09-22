VERSION 5.00
Begin VB.Form frmAcceptDeny 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Transfer Request"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton cmdDeny 
      Caption         =   "&Deny"
      Height          =   375
      Left            =   2423
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Height          =   375
      Left            =   983
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File Description"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "%1 wants to send you the file ""%2."""
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmAcceptDeny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DataArray() As String

Public Sub LoadArray(sArray() As String)
    DataArray = sArray
    Label1.Caption = DataArray(0) & " wants to send you the file """ & DataArray(1) & """. Will you accept it?"
    txtDescription.Text = DataArray(3)
End Sub

Private Sub cmdAccept_Click()
    'create a File Transfer Form in download mode
    frmMain.NewTransfer DataArray(1), CLng(DataArray(2)), DataArray(4), CLng(DataArray(5)), FTR_DNLOAD
    Unload Me
End Sub

Private Sub cmdDeny_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If DataArray(0) = "@Server" Then
        'AutoAccept!
        frmMain.NewTransfer DataArray(1), CLng(DataArray(2)), DataArray(4), CLng(DataArray(5)), FTR_DNLOAD
        Unload Me
    End If
End Sub


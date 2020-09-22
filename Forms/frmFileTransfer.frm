VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFileTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileShare in Progress - 0%"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmFileTransfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenFolder 
      Caption         =   "Open Folder"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Downloading %1 to %2"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This form maximizes efficiency  by combining both
' server and client functionality of filesharing.

' The server (sender) starts by sending a FTR to the intended client.
' the server looks for a free port, and sends that port number, file name, ip add
' and file size to the client, and waits for an inbound connection

' the client recieves the FTR, that has data regarding the port and address
' of the server, as well as the filename and size.
' the client connects to the server, whereupon the server loads the entire file
' to memory and sends it to the client.

' these local variables store information about the sender and the file
Dim mFilename As String
Dim mRemote As String
Dim mPort As Long
Dim mMode As FILETRANSFERMODE
Dim mFileSize As Long

'this tracks where we should store the data
Dim mFilePtr As Long

Dim hUpFile As Long
Dim hDnFile As Long

'this stores the actual downloaded data
Dim filedata() As Byte

'used to copy memory between the downloaded data and filedata
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'these Properties are used to set local variables from frmMain
Property Let FileSize(lFSize As Long)
    mFileSize = lFSize
End Property

Property Let FileName(sFName As String)
    mFilename = sFName
End Property

Property Let Remote(sName As String)
    mRemote = sName
End Property

Property Let Port(lPort As Long)
    mPort = lPort
End Property

'this is used to set the port value in frmMain
'when the reqquested port is in use
Property Get Port() As Long
    Port = mPort
End Property

'modes: 0 = uploading, 1 = downloading
Property Let mode(iMode As Integer)
    mMode = iMode
End Property

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub StartTransfer()
    'make sure Winsock is ready
    Winsock1.Close
    
    Select Case mMode
    Case FTR_UPLOAD
        'Uploading
        Dim mFileTitle As String
        mFileTitle = Right(mFilename, Len(mFilename) - InStr(1, mFilename, "\"))
        Label1.Caption = "Sending """ & mFileTitle & """ to " & mRemote
        
        'if an error occurs...
        On Error GoTo errBind
        'try to bind the port
        Winsock1.Bind mPort, Winsock1.LocalIP
        'turn off error handling
        On Error GoTo 0
        'wait for a connection
        Winsock1.Listen
    Case FTR_DNLOAD
        'Downloading
        ReDim filedata(mFileSize)
        
        Label1.Caption = "Downloading """ & mFilename & """ to ""Downloads"""
        'try to connect to remote side
        Winsock1.Connect mRemote, mPort
    
    End Select
    
    
    Exit Sub
    
errBind:
    '... try the next port
    mPort = mPort + 1
    Resume
End Sub

Private Sub cmdOpenFolder_Click()
    Shell GetSystemFolder(CSIDL_WINDOWS) & "\explorer.exe " & DownloadsFolder, vbNormalFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If cmdCancel.Caption = "Cancel" Then
        If MsgBox("Are you sure you want to cancel FileShare?", vbExclamation + vbYesNo + vbDefaultButton2, "Cancel FileShare") = vbNo Then
            Cancel = 1
            Exit Sub
        End If
        Winsock1.Close
    End If
End Sub

Private Sub Winsock1_Close()
Dim targetfile As String

    If mFilePtr < mFileSize Then
        Me.Caption = "FileSharing - Incomplete"
        Label1.Caption = "Download cancelled by remote side."
        cmdCancel.Caption = "Close"
        Winsock1.Close
        
        Erase filedata
        Exit Sub
    End If
    
    Me.Caption = "FileSharing - 100% complete"
    Label1.Caption = "Download complete."
    cmdCancel.Caption = "Close"

    MakeSureDirectoryExists DownloadsFolder
    
    targetfile = DownloadsFolder & "\" & mFilename
    If Dir(targetfile) <> "" Then Kill targetfile
    
    hDnFile = FreeFile
    Open targetfile For Binary As hDnFile
    Put #1, , filedata
    Close hDnFile
    
    cmdOpenFolder.Enabled = True
    
    Winsock1.Close
    
    Erase filedata
End Sub

Private Sub Winsock1_Connect()
    Winsock1.SendData "OK!"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Dim data() As Byte
    Winsock1.Close
    Winsock1.Accept requestID
    
    hUpFile = FreeFile
    Open mFilename For Binary As hUpFile
    ReDim data(LOF(1))
    Get #1, , data
    Close hUpFile
    
    Winsock1.SendData data
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim targetfile As String
Dim data() As Byte
Dim percentage As Integer

    If bytesTotal <= 1 Then Exit Sub
    
    Winsock1.GetData data, , bytesTotal
    CopyMemory filedata(mFilePtr), data(0), bytesTotal
    mFilePtr = mFilePtr + bytesTotal
    
    percentage = Int((mFilePtr / mFileSize) * 100)
    Me.Caption = "FileSharing - " & percentage & "% complete"
    ProgressBar1.Value = percentage
End Sub

Private Sub Winsock1_SendComplete()
    Select Case mMode
    Case FTR_UPLOAD
        Winsock1.Close
        Me.Caption = "FileSharing - 100% complete"
        Label1.Caption = "Upload complete."
        cmdCancel.Caption = "Close"
    Case FTR_DNLOAD
    '
    End Select
End Sub

Private Sub Winsock1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Dim percentage As Integer
    percentage = Int((mFileSize - bytesRemaining) / mFileSize * 100)
    ProgressBar1.Value = percentage
    Caption = "FileSharing - " & percentage & "% complete"
End Sub

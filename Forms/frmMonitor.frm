VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   AutoRedraw      =   -1  'True
   Caption         =   "Monitoring"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7620
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5400
   ScaleWidth      =   7620
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   120
         Top             =   3840
      End
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myDIB As cDIBSection
Dim bDIBexists As Boolean
Dim offset As Long
Dim lHeight As Long
Dim lWidth As Long
Dim picData() As Byte
Dim totalSize As Long
Dim compSize As Long

Private Const SRCPAINT = &HEE0086       ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6         ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020        ' (DWORD) dest = source

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Sub Form_Load()
    Set myDIB = New cDIBSection
End Sub
    
Public Sub SetDIBSize(dWidth As Long, dHeight As Long, CompressedSize As Long)
    
    If myDIB Is Nothing Then Exit Sub
    
    lHeight = dHeight
    lWidth = dWidth
    
    totalSize = lWidth * lHeight * 3
    
    compSize = CompressedSize
    ReDim picData(CompressedSize)
    
    If Not bDIBexists Then
        bDIBexists = myDIB.Create(dWidth, dHeight)
    End If

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Height < 28 * 15 Then Exit Sub
    If Me.Width < 8 * 15 Then Exit Sub
    Picture1.Width = Me.Width - 8 * 15
    Picture1.Height = Me.Height - 34 * 15
    If bDIBexists Then
        StretchBlt Picture1.hDC, 0, 0, Picture1.Width / Screen.TwipsPerPixelX, Picture1.Height / Screen.TwipsPerPixelY, myDIB.hDC, 0, 0, myDIB.Width, myDIB.Height, SRCCOPY
        Picture1.Refresh
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    myDIB.ClearUp
    Set myDIB = Nothing
    Erase picData
    bDIBexists = False
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If Winsock1.State = sckClosed Then Winsock1.Connect
End Sub

Private Sub Winsock1_Close()
    If Me.WindowState = vbMinimized Then Exit Sub
    If bDIBexists Then
       UpdateDisplay
       mdiServer.MonitorUser
    End If
End Sub

Public Sub UpdateDisplay()
Dim ret As Long
    Dim cComp As New cCompression
    ret = cComp.DecompressByteArray(picData, totalSize)
    If ret = Z_OK Then
        myDIB.SetByteArray 0, picData
        StretchBlt Picture1.hDC, 0, 0, Picture1.Width / Screen.TwipsPerPixelX, Picture1.Height / Screen.TwipsPerPixelY, myDIB.hDC, 0, 0, myDIB.Width, myDIB.Height, SRCCOPY
        Picture1.Refresh
    Else
        ReDim picData(totalSize)
        myDIB.SetByteArray 0, picData
        StretchBlt Picture1.hDC, 0, 0, Picture1.Width / Screen.TwipsPerPixelX, Picture1.Height / Screen.TwipsPerPixelY, myDIB.hDC, 0, 0, myDIB.Width, myDIB.Height, SRCCOPY
        Picture1.Refresh
        Debug.Print "Error decompressing - " & ret
    End If
    offset = 0
    Set cComp = Nothing
End Sub

Private Sub Winsock1_Connect()
    offset = 0
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim data() As Byte
Dim cmd As String
Dim dWidth As Long
Dim dHeight As Long
    
    If Winsock1.State = sckConnected Then
        Winsock1.GetData data, , bytesTotal
        CopyMemory picData(offset), data(0), UBound(data) + 1
        offset = offset + UBound(data) + 1
    End If

End Sub


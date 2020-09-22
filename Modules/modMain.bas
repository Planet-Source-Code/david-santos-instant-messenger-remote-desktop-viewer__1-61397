Attribute VB_Name = "modMain"
Option Explicit

Public UserName As String
Public Password As String
Public UserIndex As Integer

Public baseFileTransferPort As Long

Public mColor As Long
Public mFontName As String
Public mFontSize As Integer

Public BufferFolder As String
Public DownloadsFolder As String
Public LogsFolder As String

Public Enum FILETRANSFERMODE
    FTR_UPLOAD = 0
    FTR_DNLOAD = 1
End Enum

'used to pause for just a bit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub MakeSureDirectoryExists(path As String)
    If Dir(path, vbDirectory) = "" Then MkDir path
End Sub

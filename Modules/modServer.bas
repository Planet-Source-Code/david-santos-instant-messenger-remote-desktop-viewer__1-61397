Attribute VB_Name = "modServer"
Option Explicit
Public mColor As Long
Public mFontName As String
Public mFontSize As Integer

'in-memory record of logged users
Public Type UserRecord
    IPAddress As String
    Port As String
    UserName As String
    Password As String
    UserID As Long
    Timeout As Long
End Type

Public User() As UserRecord
'used to keep track of how many users are online
Public lastUser As Integer


Public Const Z_OK = 0
Public Const Z_STREAM_END = 1
Public Const Z_NEED_DICT = 2
Public Const Z_ERRNO = (-1)
Public Const Z_STREAM_ERROR = (-2)
Public Const Z_DATA_ERROR = (-3)
Public Const Z_MEM_ERROR = (-4)
Public Const Z_BUF_ERROR = (-5)
Public Const Z_VERSION_ERROR = (-6)


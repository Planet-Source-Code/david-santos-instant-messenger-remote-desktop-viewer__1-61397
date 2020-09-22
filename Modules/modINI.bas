Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadINI(INIFile As String, section As String, value As String, default As String) As String
    Dim tempstr As String
    Dim slength As Long  ' receives length of the returned string
    tempstr = Space(255) ' provide enough room for the function to put the value into the buffer
    slength = GetPrivateProfileString(section, value, default, tempstr, 255, INIFile)
    ReadINI = Left(tempstr, slength) ' extract the returned string from the buffer
End Function

Public Sub WriteINI(INIFile As String, section As String, keyname As String, value As String)
    Dim retval As Long
    retval = WritePrivateProfileString(section, keyname, value, INIFile)
End Sub




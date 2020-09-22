Attribute VB_Name = "modCRC32"
Option Explicit

' -------- CRC32 Module -----------
' Insert this in your project and
' use sCRC = GetCRC(filename) where
' sCRC is a string.

Private Declare Sub crc32 Lib "CRC.dll" (ByVal sMyString As String, ByVal sCRC As String)

Public Function GetCRC(fileName As String) As String
Dim mCRC As String
    mCRC = String(8, "0")
    crc32 fileName, mCRC
    GetCRC = mCRC
End Function

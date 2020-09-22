Attribute VB_Name = "modSound"
Option Explicit

Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC = &H1               '  play asynchronously
Private Const SND_FILENAME = &H20000        '  name is a file name
Private Const SND_NOWAIT = &H2000           '  don't wait if the driver is busy
Private Const SND_PURGE = &H40              '  purge non-static events for task

Public Sub PlayWav(FileName As String)
    PlaySound FileName, 0&, SND_FILENAME + SND_ASYNC + SND_NOWAIT
End Sub

Public Sub StopSound()
    PlaySound "", 0&, SND_PURGE
End Sub


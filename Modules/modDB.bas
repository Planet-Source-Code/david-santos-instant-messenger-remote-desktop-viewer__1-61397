Attribute VB_Name = "modDB"
Option Explicit

Public myConn As ADODB.Connection
Public UserRS As ADODB.Recordset

Public Sub OpenDB(fileName As String)
On Error GoTo errOpen
    
    Set myConn = New ADODB.Connection
    
    With myConn
        'setup connection properties
        'OLEDB 4.0 because we're using Access 2000 MDB
        'we could also use an ODBC driver
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = fileName
        
        '.Provider = "Microsoft.Jet.OLEDB.4.0"
        '.ConnectionString = "yapper"
        
        .Open
    End With
    
    Exit Sub

errOpen:
    MsgBox "There was an error opening the database.  The application will now terminate." & vbCrLf & vbCrLf & "Error " & Err.Number & " - " & Err.Description, vbExclamation, "Error"
    Set myConn = Nothing
    End
End Sub

' make a shortcut function so we don't have to specify all the
' options everytime we want to query the database
Public Function QueryDB(myQuery As String) As ADODB.Recordset
    Set QueryDB = New ADODB.Recordset
    QueryDB.Open myQuery, myConn, adOpenKeyset, adLockOptimistic, adCmdText
End Function

Public Sub CloseDB()
    myConn.Close
    Set myConn = Nothing
End Sub

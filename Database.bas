Attribute VB_Name = "Database"
Option Explicit

Sub ConnectToDatabase( _
    ByRef robjConnection As ADODB.Connection)
'**************************************
'purpose:connect to db
'inputs:objConnection--connection to
'       use
'returns:NONE
'explanation:NONE
'**************************************
    
    If gobjConnection Is Nothing Then
        Set gobjConnection = New ADODB.Connection
    
        On Error GoTo lblConnectError
        gobjConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "/" & "WebAgent.mdb;Persist Security Info=False", _
                , , adAsyncConnect
        On Error GoTo 0
        
        'loop while connecting
        Do While (gobjConnection.State And adStateConnecting)
            DoEvents
        Loop
    
    End If

    Set robjConnection = gobjConnection
    
Exit Sub
lblConnectError:
    Err.Raise Err.Number, Err.Source, Err.Description, _
        Err.HelpFile, Err.HelpContext
        
End Sub

Sub DisconnectFromDatabase( _
    ByRef robjConnection As ADODB.Connection)
'**************************************
'purpose:disconnect from db
'inputs:objConnection--connection to
'       use
'returns:NONE
'explanation:NONE
'**************************************

    'close connection
    'robjConnection.Close
    'Set robjConnection = Nothing

End Sub

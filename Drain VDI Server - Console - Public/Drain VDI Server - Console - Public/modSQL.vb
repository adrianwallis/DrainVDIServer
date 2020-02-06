
Imports System.Data
Module modSQL
    Public myDataReader As System.Data.SqlClient.SqlDataReader

    Dim oSQLConn As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection()
    Public Function SafeDataReader(ByVal DataReader As SqlClient.SqlDataReader) As Boolean
        SafeDataReader = False
        If Not DataReader Is Nothing Then
            If DataReader.IsClosed = False Then
                If DataReader.HasRows = True Then
                    Return True
                End If
            End If
        End If
    End Function
    Public Sub SingleCommand(ByVal SQL As String, ConnectString As String)
RetryMe:
        Dim CommandConn As SqlClient.SqlConnection
        Dim cmd As SqlClient.SqlCommand = Nothing
        CommandConn = New SqlClient.SqlConnection(ConnectString)
        Try
            CommandConn.Open()
            Do Until CommandConn.State = ConnectionState.Open
                If CommandConn.State = ConnectionState.Closed Then Console.WriteLine("Closed?") : CommandConn.ConnectionString = ConnectString
                CommandConn.Open()
            Loop
            cmd = New SqlClient.SqlCommand(SQL, CommandConn)
            cmd.CommandType = CommandType.Text
            If CommandConn.State <> ConnectionState.Open Then
                CommandConn.Open()
                Do Until CommandConn.State = ConnectionState.Open
                Loop
            End If
            cmd.ExecuteNonQuery()
            cmd.Dispose()
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Select Case Err.Description
                Case "The connection was not closed. The connection's current state is connecting."
                    GoTo RetryMe
                Case "Object reference not set to an instance of an object."
                    GoTo RetryMe
                Case "Internal connection fatal error."
                    GoTo RetryMe
                Case "The ConnectionString property has not been initialized."
                    GoTo RetryMe
                Case "Invalid operation. The connection is closed."
                    GoTo RetryMe
                Case "ExecuteNonQuery requires an open and available Connection. The connection's current state is open."
                    GoTo RetryMe
                Case "Not allowed to change the 'ConnectionString' property. The connection's current state is open."
                    GoTo RetryMe
                Case "ExecuteNonQuery requires an open and available Connection. The connection's current state is closed."
                    GoTo RetryMe
                Case "A transport-level error has occurred during connection clean-up. (provider: TCP Provider, error: 0 - The specified network name is no longer available.)Timeout expired.  The timeout period elapsed prior to completion of the operation or the server is not responding."
                    GoTo RetryMe
                Case "Timeout expired.  The timeout period elapsed prior to completion of the operation or the server is not responding."
                    GoTo RetryMe
                Case Else
                    If InStr(Err.Description, "Violation of PRIMARY KEY ") > 0 Then
                    Else
                        Console.WriteLine(Err.Description)
                        'sql error code needed here
                        Console.WriteLine(Err.Description)
                        Console.WriteLine(SQL)
                    End If
            End Select
        Finally
            If Not CommandConn Is Nothing Then
                If CommandConn.State <> ConnectionState.Closed Then CommandConn.Close()
            End If
            If Not CommandConn Is Nothing Then CommandConn.Dispose()
            CommandConn = Nothing
            If Not cmd Is Nothing Then cmd.Dispose()
            cmd = Nothing
        End Try
    End Sub
    Public Function GetSQLDataReader(ByVal SQL As String, ConnectString As String) As SqlClient.SqlDataReader
        GetSQLDataReader = Nothing
RetryMe:
        Dim GetSQLConnection As System.Data.SqlClient.SqlConnection = New System.Data.SqlClient.SqlConnection
        GetSQLConnection.ConnectionString = ConnectString
        Try
            GetSQLConnection.Open()
            Do Until GetSQLConnection.State = ConnectionState.Open

            Loop
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Console.WriteLine(SQL)
        End Try
        Dim SQLCommand As New System.Data.SqlClient.SqlCommand(SQL, GetSQLConnection)
        Try
            GetSQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.CloseConnection)
        Catch ex As Exception
            Console.WriteLine(Err.Description)
            If Err.Description = "Invalid attempt to read when no data is present." Then GoTo RetryMe
            If Err.Description = "Internal connection fatal error." Then GoTo RetryMe
            Console.WriteLine(Err.Description)
            Console.WriteLine("ERROR: " & ex.Message)
            Console.WriteLine(SQL)
        End Try
    End Function
    Public Function SafeDataItem(ByVal DR As SqlClient.SqlDataReader, ByVal ItemString As String) As String
        SafeDataItem = ""
        Try
            If Not DR.Item(ItemString) Is DBNull.Value Then
                SafeDataItem = DR.Item(ItemString)
            End If
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
        End Try
    End Function
    Public Function SafeDataGUID(ByVal DR As SqlClient.SqlDataReader, ByVal ItemString As String) As Guid
        SafeDataGUID = Nothing
        Try
            If Not DR.Item(ItemString) Is DBNull.Value Then
                SafeDataGUID = DR.Item(ItemString)
            End If
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
        End Try
    End Function

End Module

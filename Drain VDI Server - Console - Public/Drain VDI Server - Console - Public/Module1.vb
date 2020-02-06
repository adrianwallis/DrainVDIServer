Module Module1
    Private VMs As List(Of String)
    Private StillUsersLoggedOn As Int16
    Private Hostname As String
    Dim Timer1 As System.Timers.Timer
    Sub Main()
        Timer1 = New System.Timers.Timer(5000)
        Try
            'Store ConnectString in Global Variable
            Connectionstring = My.Settings.Item("ConnectionString")
            'Get Local Hostname
            Hostname = System.Net.Dns.GetHostEntry("LocalHost").HostName
            '*DEBUG ONLY

            '*DEBUG ONLY
            'Move Locally hosted VM's from rds.vm to rds.vmdrain
            If DrainServer(Hostname) Then
                'restart Remote Desktop Virtualization Host Agent on server
                If StopService(Hostname, "VMHostAgent") Then
                    'wait 5 seconds
                    Threading.Thread.Sleep(5000)
                    StartService(Hostname, "VMHostAgent")
                    'Get list of VM's on server
                    VMs = New List(Of String)
                    Dim DR As SqlClient.SqlDataReader = GetSQLDataReader("Select Name from rds.VmDRAIN where ServerName='" & Hostname & "'", Connectionstring)
                    If SafeDataReader(DR) Then
                        Do While DR.Read
                            VMs.Add(SafeDataItem(DR, "Name") & "domain.local")
                        Loop
                    End If
                    'Wait until all machines are logged off 

                    AddHandler Timer1.Elapsed, AddressOf Timer1_Tick
                    Timer1.Enabled = True
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
        Finally
            'If Timer1.enabled Then
            Do
                Threading.Thread.Sleep(5000)
            Loop
            ' End If
        End Try

    End Sub
    Private Function DrainServer(ServerName As String) As Boolean
        'Move VM's on server (servername) from rds.vm to rds.vmdrain
        Dim DR As SqlClient.SqlDataReader = Nothing

        Try

            DR = GetSQLDataReader("SELECT rds.Vm.Id, rds.Vm.VmHostId, rds.Vm.Name COLLATE SQL_Latin1_General_CP1_CI_AS AS Name, rds.Vm.PoolId, rds.Vm.CreationStatus FROM rds.Server INNER JOIN rds.RoleRdvh ON rds.Server.Id = rds.RoleRdvh.ServerId INNER JOIN rds.Vm ON rds.RoleRdvh.ServerId = rds.Vm.VmHostId WHERE (rds.Server.Name COLLATE SQL_Latin1_General_CP1_CI_AS = N'" & ServerName & "')", Connectionstring)
            If SafeDataReader(DR) Then
                Do While DR.Read
                    Dim ID As String = SafeDataGUID(DR, "Id").ToString
                    Dim VMHostId As String = SafeDataItem(DR, "VmHostId")
                    Dim Name As String = SafeDataItem(DR, "Name")
                    Dim PoolId As String = SafeDataItem(DR, "PoolId")
                    Dim CreationStatus As String = SafeDataItem(DR, "CreationStatus")
                    'insert into fake backup table (VmDRAIN)
                    SingleCommand("Insert into rds.VmDRAIN ([ID],[VmHostId],[Name],[PoolId],[CreationStatus],[ServerName]) VALUES ('" & ID & "','" & VMHostId & "','" & Name & "','" & PoolId & "','" & CreationStatus & "','" & ServerName & "')", Connectionstring)
                    'delete from primary table
                    SingleCommand("Delete from rds.Vm where ID='" & ID & "'", Connectionstring)
                Loop
            End If
            Console.WriteLine("VM's removed from " & ServerName)
            Return True
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Return False
        Finally
            If Not DR Is Nothing Then DR.Close()
        End Try
    End Function

    Private Function RestoreServer(ServerName As String) As Boolean
        'Move VM's on server (servername) from rds.vm to rds.vmdrain
        Dim DR As SqlClient.SqlDataReader = Nothing
        Try
            DR = GetSQLDataReader("Select * from rds.VmDRAIN where ServerName ='" & ServerName & "'", Connectionstring)
            If SafeDataReader(DR) Then
                Do While DR.Read
                    Dim ID As String = SafeDataGUID(DR, "Id").ToString
                    Dim VMHostId As String = SafeDataItem(DR, "VmHostId")
                    Dim Name As String = SafeDataItem(DR, "Name")
                    Dim PoolId As String = SafeDataItem(DR, "PoolId")
                    Dim CreationStatus As String = SafeDataItem(DR, "CreationStatus")
                    'insert into primary table (Vm)
                    SingleCommand("Insert into rds.Vm ([ID],[VmHostId],[Name],[PoolId],[CreationStatus]) VALUES ('" & ID & "','" & VMHostId & "','" & Name & "','" & PoolId & "','" & CreationStatus & "')", Connectionstring)
                    'delete from backup table (VMDrain)
                    SingleCommand("Delete from rds.VmDRAIN where ID='" & ID & "'", Connectionstring)
                Loop
            End If
            Console.WriteLine("VM's restored to " & ServerName)
            Return True
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Return False
        Finally
            If Not DR Is Nothing Then DR.Close()
        End Try
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        'Wait until all machines are logged off 
        Dim VM As String
        'debug only

        'Dim vms As New List(Of String)

        'VMs.Add(VM)

        'debug only

        Try
            Timer1.enabled = False
            Console.WriteLine("Timer_Tick")
            StillUsersLoggedOn = 0
            For Each VM In VMs
                Console.WriteLine("Checking " & VM)
                Dim CAS As New Cassia.TerminalServicesManager
                Dim Server As Cassia.ITerminalServer = Nothing
                Try
                    Server = CAS.GetRemoteServer(VM)
                    Server.Open()
                    Dim Session As Cassia.ITerminalServicesSession
                    For Each Session In Server.GetSessions
                        Select Case Session.SessionId
                            Case 0, 1
                            Case Else
                                Console.WriteLine(Session.ClientName & "::" & Session.UserName & "::" & Session.SessionId & "::" & Session.ConnectionState.ToString)
                                Select Case Session.ConnectionState
                                    Case Cassia.ConnectionState.Disconnected
                                        Console.WriteLine("Logging off " & Session.UserName)
                                        Session.Logoff()
                                        Console.WriteLine("Waiting 30 seconds for logoff " & Session.UserName)
                                        Threading.Thread.Sleep(30000)
                                    Case Cassia.ConnectionState.Active
                                        StillUsersLoggedOn += 1
                                End Select
                        End Select
                    Next
                Catch ex As Exception
                    Console.WriteLine("Error::" & ex.Message)
                Finally
                    If Not Server Is Nothing Then Server.Close()
                    Server.Dispose()
                End Try
            Next
            Console.WriteLine("Users still logged on: " & StillUsersLoggedOn)
            If StillUsersLoggedOn > 0 Then
                'continue to wait
            Else
                Console.WriteLine("No users logged on, restoring database entries and rebooting server")
                'reset database and reboot server
                'restore locally host VM's back from rds.vmdrain to rds.vm
                If RestoreServer(Hostname) Then
                    'reboot
                    System.Diagnostics.Process.Start("shutdown", "-r -t 00")
                End If
            End If
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
        Finally
            Timer1.enabled = True
        End Try
    End Sub
End Module

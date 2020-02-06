Module modServices
    Public Function StartService(ByVal Computer As String, ByVal Service As String) As Boolean
        Dim ServiceController As ServiceProcess.ServiceController = Nothing
        Try
            ServiceController = New ServiceProcess.ServiceController
            ServiceController.MachineName = Computer
            ServiceController.ServiceName = Service
            ServiceController.Start()
            Console.WriteLine("Service Started")
            Return True
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Return False
        Finally
            If Not ServiceController Is Nothing Then ServiceController.Dispose()
        End Try
    End Function
    Public Function StopService(ByVal Computer As String, ByVal Service As String)
        Dim ServiceController As ServiceProcess.ServiceController = Nothing
        Try
            ServiceController = New ServiceProcess.ServiceController
            ServiceController.MachineName = Computer
            ServiceController.ServiceName = Service
            ServiceController.Stop()
            Console.WriteLine("Service Stopped")
            Return True
        Catch ex As Exception
            Console.WriteLine("ERROR: " & ex.Message)
            Return False
        Finally
            If Not ServiceController Is Nothing Then ServiceController.Dispose()
        End Try
    End Function
End Module

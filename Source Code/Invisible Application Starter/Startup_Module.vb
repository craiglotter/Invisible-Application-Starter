Imports System.Diagnostics
Imports System.IO


Module Startup_Module

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception
            MsgBox("An error occurred in Invisible Application Starter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Sub main()
        Dim ApplicationName As String
        ApplicationName = "Invisible Application Starter"
        Try
            Dim aModuleName As String = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
            Dim aProcName As String = System.IO.Path.GetFileNameWithoutExtension(aModuleName)

            If Process.GetProcessesByName(aProcName).Length > 1 Then
                Application.Exit()
            Else
                Dim monitor As New Main_Screen
                monitor.ShowDialog()
                Application.Exit()
            End If
        Catch ex As Exception
            Error_Handler(ex, "Startup Module")
        End Try
    End Sub
End Module

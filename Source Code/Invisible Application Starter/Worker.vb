Imports System.IO


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public Event WorkerError(ByVal Message As Exception, ByVal identifier As String)
    Public Event WorkerComplete(ByVal queue As Integer)


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region



    Private Sub Error_Handler(ByVal message As Exception, ByVal identifier As String)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message, identifier)
            End If
        Catch ex As Exception
            MsgBox("An error occurred in Invisible Application Starter's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Activity_Logger(ByVal message As String)
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Activity Logs\" & Format(Now(), "yyyyMMdd") & "_Activity_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & message)
            filewriter.Flush()
            filewriter.Close()
        Catch ex As Exception
            Error_Handler(ex, "Activity Logger")
        End Try
    End Sub


    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    WorkerThread.Start()
                    
            End Select
        Catch ex As Exception
            Error_Handler(ex, "Choose Threads")
        End Try
    End Sub

    Private Sub WorkerExecute()
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\Monitor").Replace("\\", "\"))
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim finfo As FileInfo
            For Each finfo In dir.GetFiles()
                If finfo.Exists = True Then
                    Dim filereader As StreamReader = New StreamReader(finfo.FullName)
                    Dim read As String = ""
                    While filereader.Peek > -1
                        read = filereader.ReadLine
                        If Not read Is Nothing Then
                            Activity_Logger("Launch Attempt: " & read)
                            Try
                                ApplicationLauncher(read)
                            Catch ex As Exception
                                Error_Handler(ex, "Launch Attempt")
                            End Try
                        End If
                    End While
                    filereader.Close()
                    finfo.Delete()

                End If
                finfo = Nothing
            Next
            dir = Nothing
        Catch ex As Exception
            Error_Handler(ex, "Worker Execute")
        Finally
            RaiseEvent WorkerComplete(1)
        End Try
    End Sub

    Private Function ApplicationLauncher(ByVal apptoRun As String) As String
        Dim sresult As String = ""
        Try
            Dim myProcess As Process = New Process

            Dim executable, arguments As String
            Dim str As String()
            executable = ""
            arguments = ""
            If apptoRun.StartsWith("""") = True Then
                Dim endpos As Integer = apptoRun.IndexOf("""", apptoRun.IndexOf("""", 0) + 1)
                executable = apptoRun.Substring(0, endpos + 1)
                If apptoRun.Length >= (endpos + 3) Then
                    arguments = apptoRun.Substring(endpos + 2)
                End If
            Else
                str = apptoRun.Split(" ")
                For i As Integer = 0 To str.Length - 1
                    If i = 0 Then
                        executable = str(i)
                    Else
                        arguments = arguments & str(i) & " "
                    End If
                Next
                arguments = arguments.Remove(arguments.Length - 1, 1)
            End If
            Activity_Logger("LAUNCH ATTEMPT INITIATED")

            myProcess.StartInfo.FileName = executable.Replace("""", "")
            myProcess.StartInfo.Arguments = arguments
            Activity_Logger("Executable: " & myProcess.StartInfo.FileName)
            Activity_Logger("Arguments: " & myProcess.StartInfo.Arguments)
            myProcess.StartInfo.UseShellExecute = True

            myProcess.StartInfo.CreateNoWindow = False
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal

            myProcess.StartInfo.RedirectStandardInput = False
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.StartInfo.RedirectStandardError = False
            myProcess.Start()
            sresult = "Success"
            Return sresult

        Catch ex As Exception
            Error_Handler(ex, "Executing: " & apptoRun)
            sresult = "Fail"
        End Try
        Return sresult
    End Function

    'Private Sub ApplicationLauncher(ByVal apptoRun As String)
    '    Try
    '        Dim myProcess As Process = New Process

    '        myProcess.StartInfo.FileName = apptoRun
    '        'myProcess.StartInfo.Arguments = apptoRun
    '        myProcess.StartInfo.UseShellExecute = True

    '        myProcess.StartInfo.CreateNoWindow = False
    '        myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal

    '        myProcess.StartInfo.RedirectStandardInput = False
    '        myProcess.StartInfo.RedirectStandardOutput = False
    '        myProcess.StartInfo.RedirectStandardError = False
    '        myProcess.Start()
    '        Exit Sub

    '    Catch ex As Exception
    '        Error_Handler(ex, "Application Launcher")
    '    End Try
    'End Sub


End Class

Imports System.IO
Imports Microsoft.Win32


Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker

    Private workerbusy As Boolean = False

    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception, ByVal identifier As String)

    Public dataloaded As Boolean = False


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
    End Sub



    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 60000
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(162, 8)
        Me.ControlBox = False
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Main_Screen"
        Me.Opacity = 0
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
        Me.Text = "Invisible Application Starter"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized

    End Sub

#End Region



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







    Public Sub WorkerErrorHandler(ByVal Message As Exception, ByVal identifier As String)
        Try
            Error_Handler(Message, "Worker: " & identifier)
        Catch ex As Exception
            Error_Handler(ex, "Worker Error Handler")
        End Try
    End Sub

    Public Sub WorkerCompleteHandler(ByVal queue As Integer)
        Try
            workerbusy = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_threads()
            Worker1.Dispose()
        Catch ex As Exception
            Error_Handler(ex, "Main Screen Close")
        End Try

    End Sub


    Private Sub exit_threads()
        Try

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("Suspended") > -1 Or Worker1.WorkerThread.ThreadState.ToString.IndexOf("SuspendRequested") > -1 Then
                Worker1.WorkerThread.Resume()
            End If

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("WaitSleepJoin") > -1 Then
                Worker1.WorkerThread.Interrupt()
            End If

            If Worker1.WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1 Then
                Worker1.WorkerThread.ResetAbort()
            End If

            Worker1.WorkerThread.Abort()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim str As String
        Dim keyflag1 As Boolean = False
        Dim oReg As RegistryKey = Registry.LocalMachine


        Dim oKey As RegistryKey
        oKey = oReg
        Dim subs() As String = ("Software\Microsoft\Windows\CurrentVersion\Run").Split("\")
        For Each stri As String In subs
            oKey = oKey.OpenSubKey(stri, True)
        Next

        If Not oKey Is Nothing Then
            str = oKey.GetValue("InvisibleAppStarter")
            If Not IsNothing(str) And Not (str = "") Then
            Else
                oKey.SetValue("InvisibleAppStarter", """" & (Application.StartupPath & "\Application_Loader.exe").Replace("\\", "\") & """")
            End If
        End If

        oReg = Registry.LocalMachine
        Dim keys() As String = oReg.GetSubKeyNames()
        System.Array.Sort(keys)
        For Each str In keys
            If str.Equals("Software\Invisible Application Starter") = True Then
                keyflag1 = True
                Exit For
            End If
        Next str

        If keyflag1 = False Then
            oReg.CreateSubKey("Software\Invisible Application Starter")

        End If


        oKey = oReg
        subs = ("Software\Invisible Application Starter").Split("\")
        For Each stri As String In subs
            oKey = oKey.OpenSubKey(stri, True)
        Next

        If Not oKey Is Nothing Then
            oKey.SetValue("MonitorDirectory", """" & (Application.StartupPath & "\Monitor").Replace("\\", "\") & """")
        End If
        oKey.Close()
        oReg.Close()

        Timer1.Start()
        If workerbusy = False Then
            workerbusy = True
            Worker1.ChooseThreads(1)
        End If
        dataloaded = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If workerbusy = False Then
            workerbusy = True
            Worker1.ChooseThreads(1)
        End If
    End Sub
End Class

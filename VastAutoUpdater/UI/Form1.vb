Imports MaterialSkin
Imports MaterialSkin.Controls

Public Class VASTUpdater
    Inherits MaterialForm

    Private engine As New UpdaterEngine()

    Private txtSftpUsername As MaterialTextBox
    Private txtSftpPassword As MaterialTextBox
    Private btnCheckForUpdates As MaterialButton
    Private progressBar1 As MaterialProgressBar
    Private runSilently As Boolean

    Public Sub New()
        InitializeComponent()
        InitializeUX()

        Dim args As String() = Environment.GetCommandLineArgs()
        runSilently = args.Contains("silent")

        AddHandler Me.Load, AddressOf VASTUpdater_Load

        If runSilently Then
            ' Hide the form entirely in silent mode — no UI flash
            Me.WindowState = FormWindowState.Minimized
            Me.ShowInTaskbar = False
            Me.Opacity = 0
            Logger.Log("Starting in silent mode", Logger.LogLevel.Info)
            RunSilentUpdate()
        End If
    End Sub

    Private Sub VASTUpdater_Load(sender As Object, e As EventArgs)
        Try
            Dim exePath As String = VersionService.FindVastExecutable()
            If Not String.IsNullOrEmpty(exePath) Then
                Dim version As String = VersionService.GetFileVersion(exePath)
                lblCurrentVersion.Text = $"Current Version: {version}"
            Else
                lblCurrentVersion.Text = "Current Version: Not Found"
            End If
            lblStatus.Text = "Status: Ready for update check..."
        Catch ex As Exception
            Logger.Log($"Error on load: {ex.Message}", Logger.LogLevel.Error)
            lblStatus.Text = "Status: Error during load"
        End Try
    End Sub

    Private Sub InitializeUX()
        Me.Text = "VAST Updater"
        Me.Size = New Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        Dim centerX As Integer = (Me.ClientSize.Width - 300) \ 2

        txtSftpUsername = New MaterialTextBox With {
            .Hint = "SFTP Username",
            .Size = New Size(300, 50),
            .Location = New Point(centerX, 100)
        }
        Me.Controls.Add(txtSftpUsername)

        txtSftpPassword = New MaterialTextBox With {
            .Hint = "SFTP Password",
            .Password = True,
            .Size = New Size(300, 50),
            .Location = New Point(centerX, 160)
        }
        Me.Controls.Add(txtSftpPassword)

        btnCheckForUpdates = New MaterialButton With {
            .Text = "Check",
            .Size = New Size(200, 40),
            .Location = New Point(centerX + 50, 230)
        }
        AddHandler btnCheckForUpdates.Click, AddressOf btnCheckForUpdates_Click
        Me.Controls.Add(btnCheckForUpdates)

        progressBar1 = New MaterialProgressBar With {
            .Size = New Size(300, 10),
            .Location = New Point(centerX, 280),
            .Visible = False
        }
        Me.Controls.Add(progressBar1)
    End Sub

    Private Async Sub btnCheckForUpdates_Click(sender As Object, e As EventArgs)
        btnCheckForUpdates.Enabled = False
        Await RunUpdate()
        btnCheckForUpdates.Enabled = True
    End Sub

    Private Async Sub RunSilentUpdate()
        Try
            Await RunUpdate()
            Logger.Log("Silent mode finished successfully, exiting", Logger.LogLevel.Info)
            ExitApplication(0)
        Catch ex As Exception
            Logger.Log($"Silent mode failed with unhandled error: {ex.Message}", Logger.LogLevel.Error)
            ExitApplication(1)
        End Try
    End Sub

    Private Async Function RunUpdate() As Task
        Dim user As String
        Dim pass As String

        If runSilently Then
            ' In silent mode, read SFTP credentials from App.config
            user = ConfigManager.SftpUsername
            pass = ConfigManager.SftpPassword
            Logger.Log("Silent mode: using credentials from configuration", Logger.LogLevel.Info)
        Else
            user = txtSftpUsername.Text
            pass = txtSftpPassword.Text
        End If

        If String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(pass) Then
            Logger.Log("Credentials required but not available", Logger.LogLevel.Warning)
            If Not runSilently Then MessageBox.Show("Enter SFTP credentials")
            Return
        End If

        If Not runSilently Then
            progressBar1.Visible = True
            lblStatus.Text = "Starting update check..."
        End If

        Try
            Await engine.PerformUpdateCheck(user, pass,
                Sub(p As Integer, status As String)
                    If runSilently Then Return
                    Try
                        Me.Invoke(Sub()
                                      lblStatus.Text = status
                                      progressBar1.Value = Math.Min(Math.Max(p, 0), 100)
                                  End Sub)
                    Catch
                        ' Form may be disposed during exit
                    End Try
                End Sub)

            If Not runSilently Then
                Me.Invoke(Sub()
                              lblStatus.Text = "Update complete."
                              progressBar1.Visible = False
                          End Sub)
                Logger.Log("Update completed in UI mode.", Logger.LogLevel.Info)
            End If

        Catch ex As Exception
            Logger.Log($"Update failed: {ex.Message}", Logger.LogLevel.Error)

            If Not runSilently Then
                Try
                    Me.Invoke(Sub()
                                  lblStatus.Text = $"Error: {ex.Message}"
                                  progressBar1.Visible = False
                              End Sub)
                Catch
                    ' Form may be disposed
                End Try
            End If

            ' Re-throw for silent mode to catch in RunSilentUpdate
            If runSilently Then Throw
        End Try
    End Function

    ''' <summary>
    ''' Centralized exit point — ensures logging before process termination.
    ''' </summary>
    Private Sub ExitApplication(exitCode As Integer)
        Logger.Log($"Application exiting with code {exitCode}", Logger.LogLevel.Info)
        Environment.Exit(exitCode)
    End Sub

End Class

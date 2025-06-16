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

        Dim args = Environment.GetCommandLineArgs()
        runSilently = args.Contains("silent")

        AddHandler Me.Load, AddressOf VASTUpdater_Load

        If runSilently Then
            Logger.Log("Starting in silent mode", Logger.LogLevel.Info)
            RunSilentUpdate()
        End If
    End Sub

    Private Sub VASTUpdater_Load(sender As Object, e As EventArgs)
        Try
            Dim exePath = VersionService.FindVastExecutable()
            If Not String.IsNullOrEmpty(exePath) Then
                Dim version = VersionService.GetFileVersion(exePath)
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
        Await RunUpdate()

        Logger.Log("Silent mode finished, exiting", Logger.LogLevel.Info)
        Environment.Exit(0)
    End Sub

    Private Async Function RunUpdate() As Task
        Dim user = txtSftpUsername.Text
        Dim pass = txtSftpPassword.Text

        If String.IsNullOrWhiteSpace(user) OrElse String.IsNullOrWhiteSpace(pass) Then
            Logger.Log("Credentials required", Logger.LogLevel.Warning)
            If Not runSilently Then MessageBox.Show("Enter SFTP credentials")
            Return
        End If

        progressBar1.Visible = True
        lblStatus.Text = "Starting update check..."

        Try
            Await engine.PerformUpdateCheck(user, pass, Sub(p, status)
                                                            Me.Invoke(Sub()
                                                                          lblStatus.Text = status
                                                                          progressBar1.Value = Math.Min(p, 100)
                                                                      End Sub)
                                                        End Sub)

            Me.Invoke(Sub()
                          lblStatus.Text = "Update complete."
                          progressBar1.Visible = False
                      End Sub)

            If runSilently Then
                Logger.Log("Silent update successful. Exiting.", Logger.LogLevel.Info)
                Environment.Exit(0)
            Else
                Logger.Log("Update completed in UI mode. Closing.", Logger.LogLevel.Info)
                Application.Exit()
            End If

        Catch ex As Exception
            Logger.Log($"Update failed: {ex.Message}", Logger.LogLevel.Error)

            Me.Invoke(Sub()
                          lblStatus.Text = $"Error: {ex.Message}"
                          progressBar1.Visible = False

                          If runSilently Then
                              Logger.Log("Silent mode: exiting due to failure", Logger.LogLevel.Info)
                              Environment.Exit(1)
                          Else
                              Logger.Log("UI mode: closing due to fatal error", Logger.LogLevel.Info)
                              Me.Close()
                              Application.Exit()
                          End If
                      End Sub)
        End Try
    End Function
End Class

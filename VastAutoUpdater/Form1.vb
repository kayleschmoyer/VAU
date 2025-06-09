Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports MaterialSkin
Imports MaterialSkin.Controls
Imports Renci.SshNet
Imports System.Security.AccessControl
Imports System.Security.Principal

Public Class VASTUpdater
    Inherits MaterialForm

    ' SFTP Configuration
    Private ReadOnly SftpHost As String = "ftp.kerridgecsna.com"
    Private ReadOnly RemoteDirectory As String = "/VASTAutoInstall/"

    ' Email Notification Configuration (hard-coded for now)
    Private ReadOnly SmtpHost As String = "smtp.example.com"
    Private ReadOnly SmtpPort As Integer = 587
    Private ReadOnly SmtpUser As String = "noreply@yourcompany.com"
    Private ReadOnly SmtpPass As String = "YourSmtpPassword"
    Private ReadOnly EmailRecipients As String() = {"ops1@yourcompany.com", "ops2@yourcompany.com"}

    ' SQL Configuration (to look up store info)
    Private ReadOnly SqlConnectionString As String = "Server=YOUR_SQL_SERVER;Database=YOUR_DB;User Id=SQLUser;Password=SQLPass;"

    ' UI Controls
    Private txtSftpUsername As MaterialTextBox
    Private txtSftpPassword As MaterialTextBox
    Private btnCheckForUpdates As MaterialButton
    Private progressBar1 As MaterialProgressBar
    Private lblStatus As MaterialLabel
    Private lblCurrentVersion As MaterialLabel

    Private totalBytes As ULong
    Private runSilently As Boolean = False
    Private fadeTimer As System.Windows.Forms.Timer
    Private lastLoggedProgress As Integer = -1
    Private downloadComplete As Boolean = False

    Private Const LOG_BASE_PATH As String = "VASTUpdater"
    Private Const INSTALL_BASE_PATH As String = "VASTUpdater\NewPatchInstall"
    Private Const OLD_PATCHES_PATH As String = "VASTUpdater\Old Patches"

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.ClientSize = New System.Drawing.Size(600, 400)
        Me.Name = "VASTUpdater"
        Me.Text = "VAST Updater"
        Me.ResumeLayout(False)
    End Sub

    Public Sub New()
        Try
            Dim args As String() = Environment.GetCommandLineArgs()
            runSilently = args.Contains("silent")

            If runSilently Then
                LogMessage("Starting in silent mode via Task Scheduler...", "INFO")
                RunSilentUpdate().GetAwaiter().GetResult()
                LogMessage("Ensuring process exit in silent mode...", "INFO")
                Environment.Exit(0)
                Return
            End If

            LogMessage("Initializing application...", "INFO")
            ClearLogsOnStartup()

            InitializeComponent()
            InitializeUX()

            Dim skinManager = MaterialSkinManager.Instance
            skinManager.AddFormToManage(Me)
            skinManager.Theme = MaterialSkinManager.Themes.LIGHT
            skinManager.ColorScheme = New ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.Pink400, TextShade.WHITE)

            CreateScheduledTask()
        Catch ex As Exception
            LogMessage($"❌ Error during initialization: {ex.Message}", "ERROR")
            WriteDebugLog($"Initialization failed: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                MessageBox.Show($"Startup Error: {ex.Message}", "Startup Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
            LogMessage("Ensuring process exit after initialization error...", "INFO")
            Environment.Exit(1)
        End Try
    End Sub

    Private Sub ClearLogsOnStartup()
        Try
            Dim updateLogPath As String = GetLogPath()
            If Not File.Exists(updateLogPath) Then
                File.WriteAllText(updateLogPath, $"UpdateLog.txt created at {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}")
            Else
                File.WriteAllText(updateLogPath, String.Empty)
            End If

            Dim debugLogPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater", "DebugLog.txt")
            Dim debugLogDir As String = Path.GetDirectoryName(debugLogPath)
            If Not Directory.Exists(debugLogDir) Then Directory.CreateDirectory(debugLogDir)
            If Not File.Exists(debugLogPath) Then
                File.WriteAllText(debugLogPath, $"DebugLog.txt created at {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}")
            Else
                File.WriteAllText(debugLogPath, String.Empty)
            End If
        Catch ex As Exception
            WriteDebugLog($"Failed to clear logs on startup: {ex.Message}")
            Try
                Dim fallbackDebugPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VASTUpdater", "DebugLog.txt")
                Dim fallbackDir As String = Path.GetDirectoryName(fallbackDebugPath)
                If Not Directory.Exists(fallbackDir) Then Directory.CreateDirectory(fallbackDir)
                File.WriteAllText(fallbackDebugPath, $"Fallback DebugLog.txt - Failed to clear logs on startup: {ex.Message}{Environment.NewLine}")
            Catch
                ' Silent fail
            End Try
        End Try
    End Sub

    Private Sub InitializeUX()
        Me.Text = "VAST Updater"
        Me.Size = New System.Drawing.Size(600, 400)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.AutoScaleMode = AutoScaleMode.Font

        Dim centerX As Integer = (Me.ClientSize.Width - 300) \ 2

        txtSftpUsername = New MaterialTextBox With {
            .Hint = "Enter SFTP Username",
            .Size = New Size(300, 50),
            .Location = New Point(centerX, 100)
        }
        Me.Controls.Add(txtSftpUsername)

        txtSftpPassword = New MaterialTextBox With {
            .Hint = "Enter SFTP Password",
            .Password = True,
            .Size = New Size(300, 50),
            .Location = New Point(centerX, 160)
        }
        Me.Controls.Add(txtSftpPassword)

        btnCheckForUpdates = New MaterialButton With {
            .Text = "Initiate Update",
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

        lblStatus = New MaterialLabel With {
            .Text = "Status: Ready for update check...",
            .Location = New Point(centerX, 300),
            .Size = New Size(300, 30),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.WhiteSmoke
        }
        Me.Controls.Add(lblStatus)

        lblCurrentVersion = New MaterialLabel With {
            .Text = "Current Version: Checking...",
            .Location = New Point(centerX, 330),
            .Size = New Size(300, 30),
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = Color.WhiteSmoke
        }
        Me.Controls.Add(lblCurrentVersion)
    End Sub

    Private Async Function RunSilentUpdate() As Task
        Try
            LogMessage("Running silent update check...", "INFO")
            If String.IsNullOrEmpty(SftpHost) Then
                LogMessage("❌ ERROR: SFTP Host is not set.", "ERROR")
                Throw New ArgumentException("SFTP Host is not set.")
            End If

            Dim sftpUsername As String = "VastAutoInstall"
            Dim sftpPassword As String = "FwNN$ec53wBleN87W1"
            If String.IsNullOrEmpty(sftpUsername) OrElse String.IsNullOrEmpty(sftpPassword) Then
                LogMessage("❌ ERROR: SFTP credentials missing.", "ERROR")
                Throw New ArgumentException("SFTP credentials missing.")
            End If

            LogMessage($"Using SFTP credentials - Username: {sftpUsername}", "INFO")
            Await PerformUpdateCheck(sftpUsername, sftpPassword)
            LogMessage("Silent update check completed.", "INFO")
        Catch ex As Exception
            LogMessage($"❌ Silent update failed: {ex.Message}", "ERROR")
            WriteDebugLog($"Silent update failed: {ex.Message}, Stacktrace: {ex.StackTrace}")
            Throw
        Finally
            LogMessage("Silent update process completed. Exiting...", "INFO")
            Environment.Exit(0)
        End Try
    End Function

    Private Async Sub btnCheckForUpdates_Click(sender As Object, e As EventArgs)
        Try
            Dim sftpUsername As String = txtSftpUsername.Text
            Dim sftpPassword As String = txtSftpPassword.Text

            If String.IsNullOrWhiteSpace(sftpUsername) OrElse String.IsNullOrWhiteSpace(sftpPassword) Then
                LogMessage("Credentials required: Please enter your SFTP username and password.", "WARNING")
                MessageBox.Show("Credentials required: Please enter your SFTP username and password.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If Not runSilently Then
                Me.Invoke(Sub()
                              lblStatus.Text = "Status: Initiating update check..."
                              ShowLoadingOverlay("Initiating update check...")
                          End Sub)
            End If

            Await PerformUpdateCheck(sftpUsername, sftpPassword)
        Catch ex As Exception
            LogMessage($"Error during update check: {ex.Message}", "ERROR")
            WriteDebugLog($"Update check failed in silent mode: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                Me.Invoke(Sub()
                              HideLoadingOverlay()
                              lblStatus.Text = "Status: Update check failed."
                              progressBar1.Visible = False
                              MessageBox.Show($"Error during update check: {ex.Message}", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                          End Sub)
            End If
        Finally
            If runSilently Then
                LogMessage("CheckForUpdates completed. Ensuring process exit in silent mode...", "INFO")
                Environment.Exit(0)
            End If
        End Try
    End Sub

    Private Async Function PerformUpdateCheck(sftpUsername As String, sftpPassword As String) As Task
        Try
            ' 1) Locate local VAST.exe
            Dim localFilePath = FindVastExecutable()
            If String.IsNullOrEmpty(localFilePath) Then
                LogMessage("❌ VAST.exe not found.", "ERROR")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  lblStatus.Text = "Status: VAST.exe not found."
                                  MessageBox.Show("VAST.exe not found. Update check cannot proceed.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              End Sub)
                End If
                Return
            End If

            ' 2) Determine current version
            Dim currentVersionString = GetFileVersion(localFilePath)
            Dim currentVersion = New Version(currentVersionString)
            Dim majorMinorPrefix = $"{currentVersion.Major}.{currentVersion.Minor}"
            LogMessage($"Current Version: {currentVersion}", "INFO")

            ' 3) UI feedback
            If Not runSilently Then
                Me.Invoke(Sub()
                              lblStatus.Text = "Status: Checking for updates..."
                              progressBar1.Visible = True
                              progressBar1.Value = 10
                          End Sub)
            End If

            ' 4) SFTP: find latest
            Dim latestVersionString = GetLatestVersionFromSFTP(sftpUsername, sftpPassword, majorMinorPrefix)
            Dim latestVersion As Version
            If Not Version.TryParse(latestVersionString, latestVersion) Then
                latestVersion = New Version("0.0.0")
            End If

            If latestVersion.CompareTo(New Version("0.0.0")) = 0 Then
                LogMessage("❌ No valid update found.", "ERROR")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  HideLoadingOverlay()
                                  lblStatus.Text = "Status: No valid update found."
                                  MessageBox.Show("No valid update found.", "Update Check", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                  progressBar1.Visible = False
                              End Sub)
                End If
                Return
            End If

            ' 5) Compare and update
            If latestVersion.CompareTo(currentVersion) >= 0 Then
                LogMessage($"Update candidate found: {latestVersion}. Downloading...", "INFO")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  lblStatus.Text = $"Status: Downloading update {latestVersion}..."
                                  progressBar1.Value = 30
                              End Sub)
                End If

                If DownloadNewVersion(sftpUsername, sftpPassword, latestVersion.ToString()) Then
                    Dim downloadPath = GetInstallPath(latestVersion.ToString())
                    LogMessage($"Running downloaded update: {downloadPath}", "INFO")
                    If Not runSilently Then
                        Me.Invoke(Sub()
                                      lblStatus.Text = "Status: Installing update..."
                                      progressBar1.Value = 50
                                  End Sub)
                    End If

                    Dim installSuccess = Await ShowInstallProgress(downloadPath)
                    If installSuccess Then
                        LogMessage($"Successfully applied update {latestVersion}.", "INFO")
                        If Not runSilently Then
                            Me.Invoke(Sub()
                                          HideLoadingOverlay()
                                          lblStatus.Text = $"Status: Updated to version {latestVersion}"
                                          progressBar1.Visible = False
                                          progressBar1.Value = 0
                                      End Sub)
                        End If
                    Else
                        LogMessage("Update installation failed.", "ERROR")
                        If Not runSilently Then
                            Me.Invoke(Sub()
                                          HideLoadingOverlay()
                                          lblStatus.Text = "Status: Installation failed."
                                          progressBar1.Visible = False
                                          progressBar1.Value = 0
                                      End Sub)
                        End If
                    End If
                Else
                    LogMessage("Update download failed.", "ERROR")
                    If Not runSilently Then
                        Me.Invoke(Sub()
                                      HideLoadingOverlay()
                                      lblStatus.Text = "Status: Download failed."
                                      progressBar1.Visible = False
                                      MessageBox.Show("Update download failed.", "Download Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                  End Sub)
                    End If
                End If
            Else
                LogMessage("No updates available. Already up-to-date.", "INFO")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  HideLoadingOverlay()
                                  lblStatus.Text = "Status: Up-to-date."
                                  progressBar1.Visible = False
                                  MessageBox.Show("You are up-to-date!", "No Updates Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                              End Sub)
                End If
            End If

        Catch ex As Exception
            LogMessage($"Error in PerformUpdateCheck: {ex.Message}", "ERROR")
            WriteDebugLog($"PerformUpdateCheck failed: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                Me.Invoke(Sub()
                              HideLoadingOverlay()
                              lblStatus.Text = "Status: Update check failed."
                              progressBar1.Visible = False
                              MessageBox.Show($"Error during update check: {ex.Message}", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                          End Sub)
            End If
        Finally
            If runSilently Then
                LogMessage("PerformUpdateCheck completed. Exiting silent mode...", "INFO")
                Environment.Exit(0)
            End If
        End Try
    End Function

    Private Async Function ShowInstallProgress(filePath As String) As Task(Of Boolean)
        Dim installSuccess As Boolean = False
        Dim statusMessage As String = ""
        Dim silent As Boolean = Environment.GetCommandLineArgs().Contains("silent")

        Try
            If Not File.Exists(filePath) Then
                statusMessage = "Installer not found"
                LogMessage($"❌ {statusMessage}", "ERROR")
                SendNotification(False, statusMessage)
                If Not silent Then
                    Me.Invoke(Sub()
                                  HideLoadingOverlay()
                                  lblStatus.Text = "Status: Installer not found."
                                  MessageBox.Show("Update installer not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                              End Sub)
                End If
                Return False
            End If

            If Not silent Then
                Me.Invoke(Sub()
                              ShowLoadingOverlay("Preparing installation...")
                              lblStatus.Text = "Status: Preparing installation..."
                          End Sub)
            End If

            installSuccess = InstallUpdate(filePath, silent, statusMessage)

            If Not silent Then
                Me.Invoke(Sub()
                              HideLoadingOverlay()
                              If installSuccess Then
                                  lblStatus.Text = "Status: Update installed successfully."
                              Else
                                  lblStatus.Text = $"Status: {statusMessage}"
                                  If statusMessage.Contains("canceled") Then
                                      MessageBox.Show("Installation was canceled by the user.", "Installation Canceled", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                  Else
                                      MessageBox.Show(statusMessage, "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                  End If
                              End If
                          End Sub)
            End If
        Catch ex As Exception
            statusMessage = $"Failed to install: {ex.Message}"
            LogMessage($"Installation failed: {ex.Message}", "ERROR")
            WriteDebugLog($"Installation failed: {ex.Message}, Stacktrace: {ex.StackTrace}")
            SendNotification(False, statusMessage)
            If Not silent Then
                Me.Invoke(Sub()
                              HideLoadingOverlay()
                              lblStatus.Text = "Status: Installation failed."
                              MessageBox.Show($"Failed to install the update: {ex.Message}", "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                          End Sub)
            End If
            installSuccess = False
        Finally
            If silent Then
                LogMessage("Installation attempt completed. Exiting silent mode...", "INFO")
                Environment.Exit(0)
            End If
        End Try

        Return installSuccess
    End Function

    Private Function InstallUpdate(filePath As String, runSilently As Boolean, ByRef statusMessage As String) As Boolean
        Dim result As Boolean = False
        Dim localStatusMessage As String = statusMessage
        Dim process As New Process()

        Try
            If Not File.Exists(filePath) Then
                localStatusMessage = "Installer not found"
                LogMessage($"❌ {localStatusMessage} at: {filePath}", "ERROR")
                statusMessage = localStatusMessage
                SendNotification(False, localStatusMessage)
                Return False
            End If

            LogMessage($"Installer file accessible at: {filePath}", "INFO")
            Dim currentUser As String = Environment.UserName
            LogMessage($"Running as user: {currentUser} (Is SYSTEM: {currentUser.ToLower() = "system"})", "INFO")
            LogMessage($"Starting installation of {filePath} in {(If(runSilently, "silent", "interactive"))} mode...", "INFO")

            ' Close or wait for all relevant processes
            Dim procsToKill = New String() {"Vast", "Commshop", "Purchasing", "AR", "VastReportingCR", "EOD", "EOM", "VastMaint"}
            For Each pname In procsToKill
                For Each proc In Process.GetProcesses().Where(Function(p) p.ProcessName.Equals(pname, StringComparison.OrdinalIgnoreCase))
                    Try
                        If Not proc.HasExited Then
                            If pname = "Commshop" Then
                                LogMessage($"Waiting for {pname}.exe (PID: {proc.Id}) to finish...", "INFO")
                                While Not proc.HasExited
                                    Threading.Thread.Sleep(1000)
                                    LogMessage($"Still waiting for {pname}.exe to finish...", "INFO")
                                End While
                                LogMessage($"{pname}.exe (PID: {proc.Id}) has finished.", "INFO")
                            Else
                                LogMessage($"Closing {pname}.exe (PID: {proc.Id})...", "INFO")
                                proc.CloseMainWindow()
                                proc.WaitForExit(5000)
                                If Not proc.HasExited Then
                                    proc.Kill()
                                    LogMessage($"Forced termination of {pname}.exe (PID: {proc.Id})", "WARNING")
                                Else
                                    LogMessage($"{pname}.exe (PID: {proc.Id}) closed gracefully.", "INFO")
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        localStatusMessage = $"Error handling {pname}.exe: {ex.Message}"
                        LogMessage($"❌ {localStatusMessage}", "ERROR")
                        statusMessage = localStatusMessage
                        SendNotification(False, localStatusMessage)
                        Return False
                    End Try
                Next
            Next

            ' Grant SYSTEM full control to installer
            LogMessage("Granting SYSTEM full control to the installer...", "INFO")
            GrantSystemFullControl(filePath, False)

            ' Configure process start
            process.StartInfo.FileName = filePath
            process.StartInfo.WorkingDirectory = Path.GetDirectoryName(filePath)
            process.StartInfo.Arguments = If(runSilently, "/silent", "")
            process.StartInfo.UseShellExecute = False
            process.StartInfo.CreateNoWindow = runSilently
            process.StartInfo.RedirectStandardOutput = True
            process.StartInfo.RedirectStandardError = True

            LogMessage("Launching installer process with CreateProcess…", "INFO")
            Try
                process.Start()
            Catch ex As ComponentModel.Win32Exception When ex.NativeErrorCode = 193
                LogMessage("❗ ERROR_BAD_EXE_FORMAT detected; retrying with ShellExecute…", "WARNING")
                process.StartInfo.UseShellExecute = True
                process.StartInfo.RedirectStandardOutput = False
                process.StartInfo.RedirectStandardError = False
                process.Start()
            End Try

            LogMessage($"Installer process started with PID: {process.Id}", "INFO")
            Dim output As String = process.StandardOutput.ReadToEnd()
            Dim errorOutput As String = process.StandardError.ReadToEnd()
            LogMessage($"Installer output: {output}", "INFO")
            If Not String.IsNullOrEmpty(errorOutput) Then LogMessage($"Installer error: {errorOutput}", "ERROR")

            LogMessage("Waiting for installer process to exit (60-second timeout)...", "INFO")
            If Not process.WaitForExit(60000) Then
                process.Kill()
                localStatusMessage = "Installation timed out"
                LogMessage($"❌ {localStatusMessage}", "ERROR")
                statusMessage = localStatusMessage
                SendNotification(False, localStatusMessage)
                Return False
            End If

            Dim exitCode As Integer = process.ExitCode
            LogMessage($"Installer process exited with code: {exitCode}", "INFO")

            If exitCode = 0 Then
                localStatusMessage = "Installation successful"
                LogMessage("Installation completed successfully.", "INFO")

                ' Move old installer to Old Patches
                Dim oldPatchesPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), OLD_PATCHES_PATH)
                If Not Directory.Exists(oldPatchesPath) Then
                    Directory.CreateDirectory(oldPatchesPath)
                    LogMessage($"Created 'Old Patches' folder at: {oldPatchesPath}", "INFO")
                End If
                Dim destinationPath As String = Path.Combine(oldPatchesPath, Path.GetFileName(filePath))
                If File.Exists(filePath) Then
                    File.Move(filePath, destinationPath)
                    LogMessage($"Moved old patch to: {destinationPath}", "INFO")
                Else
                    LogMessage($"❌ Failed to move old patch: Source file {filePath} not found.", "ERROR")
                End If

                statusMessage = localStatusMessage
                SendNotification(True, localStatusMessage)
                Return True

            ElseIf exitCode = 1223 Then
                localStatusMessage = "Installation canceled by user"
                LogMessage($"{localStatusMessage}.", "WARNING")
                statusMessage = localStatusMessage
                SendNotification(False, localStatusMessage)
                Return False

            Else
                localStatusMessage = $"Installation failed with exit code {exitCode}"
                LogMessage($"{localStatusMessage}. Check log: {GetLogPath()}", "WARNING")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  MessageBox.Show($"{localStatusMessage}. You may need to run as Administrator.", "Installation Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                              End Sub)
                End If
                statusMessage = localStatusMessage
                SendNotification(False, localStatusMessage)
                Return False
            End If

        Catch ex As Exception
            localStatusMessage = $"Failed to install update: {ex.Message}"
            LogMessage($"❌ {localStatusMessage}", "ERROR")
            LogMessage($"Exception details: {ex}", "ERROR")
            WriteDebugLog($"Install update failed: {ex.Message}, Stacktrace: {ex.StackTrace}")
            statusMessage = localStatusMessage
            SendNotification(False, localStatusMessage)
            Return False
        Finally
            If runSilently Then
                LogMessage("Installation attempt completed. Exiting silent mode...", "INFO")
                Environment.Exit(0)
            End If
        End Try
    End Function

    ''' <summary>
    ''' Send notification email on success or failure,
    ''' including store number/name and machine name.
    ''' </summary>
    Private Sub SendNotification(success As Boolean, details As String)
        Try
            ' Lookup company/store info
            Dim companyNumber As String = ""
            Dim companyName As String = ""
            Using conn As New SqlConnection(SqlConnectionString)
                conn.Open()
                Using cmd As New SqlCommand("SELECT TOP 1 COMPANY_NUMBER, NAME FROM COMPANY", conn)
                    Using rdr = cmd.ExecuteReader()
                        If rdr.Read() Then
                            companyNumber = rdr("COMPANY_NUMBER").ToString()
                            companyName = rdr("NAME").ToString()
                        End If
                    End Using
                End Using
            End Using

            Dim computer As String = Environment.MachineName
            Dim statusText As String = If(success, "Success", "Failure")
            Dim subject As String = $"Store {companyNumber}-{companyName} on {computer}: Update {statusText}"
            Dim body As New Text.StringBuilder()
            body.AppendLine($"Update {statusText} at store {companyNumber} - {companyName}")
            body.AppendLine($"Computer: {computer}")
            body.AppendLine($"Details: {details}")
            body.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")

            Using msg As New MailMessage()
                msg.From = New MailAddress(SmtpUser)
                For Each recip In EmailRecipients
                    msg.To.Add(recip)
                Next
                msg.Subject = subject
                msg.Body = body.ToString()
                msg.IsBodyHtml = False

                Using client As New SmtpClient(SmtpHost, SmtpPort)
                    client.Credentials = New Net.NetworkCredential(SmtpUser, SmtpPass)
                    client.EnableSsl = True
                    client.Send(msg)
                End Using
            End Using

            LogMessage($"Notification email sent: {subject}", "INFO")
        Catch ex As Exception
            WriteDebugLog($"Failed to send notification email: {ex.Message}")
        End Try
    End Sub

    Private Sub ShowLoadingOverlay(text As String)
        If runSilently Then Return
        Me.Invoke(Sub()
                      lblStatus.Text = text
                      progressBar1.Visible = True
                      progressBar1.Value = 0
                  End Sub)

        If fadeTimer Is Nothing Then
            fadeTimer = New System.Windows.Forms.Timer With {.Interval = 50}
            AddHandler fadeTimer.Tick, AddressOf FadeInTick
        End If
        fadeTimer.Tag = progressBar1
        fadeTimer.Start()
    End Sub

    Private Sub FadeInTick(sender As Object, e As EventArgs)
        Dim timer As System.Windows.Forms.Timer = DirectCast(sender, System.Windows.Forms.Timer)
        Dim progressBar As MaterialProgressBar = DirectCast(timer.Tag, MaterialProgressBar)
        If progressBar.Value < 10 Then
            progressBar.Value += 1
        Else
            timer.Stop()
        End If
    End Sub

    Private Sub HideLoadingOverlay()
        If runSilently OrElse progressBar1 Is Nothing Then Return
        If fadeTimer Is Nothing Then
            fadeTimer = New System.Windows.Forms.Timer With {.Interval = 50}
            AddHandler fadeTimer.Tick, AddressOf FadeOutTick
        End If
        fadeTimer.Tag = progressBar1
        fadeTimer.Start()
    End Sub

    Private Sub FadeOutTick(sender As Object, e As EventArgs)
        Dim timer As System.Windows.Forms.Timer = DirectCast(sender, System.Windows.Forms.Timer)
        Dim progressBar As MaterialProgressBar = DirectCast(timer.Tag, MaterialProgressBar)
        If progressBar.Value > 0 Then
            progressBar.Value -= 1
        Else
            timer.Stop()
            Me.Invoke(Sub()
                          progressBar1.Visible = False
                          progressBar1.Value = 0
                      End Sub)
        End If
    End Sub

    Private Function DownloadNewVersion(username As String, password As String, version As String) As Boolean
        Try
            downloadComplete = False
            EnsureUpdateFolderExists()
            Dim localPath As String = GetInstallPath(version)
            LogMessage($"Starting download: {localPath}", "INFO")

            Using client As New SftpClient(SftpHost, username, password)
                client.Connect()
                Dim remotePath As String = $"{RemoteDirectory}{version}.exe"
                LogMessage($"Downloading update from: {remotePath}", "INFO")

                If File.Exists(localPath) Then
                    File.Delete(localPath)
                    LogMessage($"Deleted old update file: {localPath}", "INFO")
                End If

                Using fileStream As New FileStream(localPath, FileMode.Create, FileAccess.Write)
                    Dim fileInfo = client.Get(remotePath)
                    totalBytes = CULng(fileInfo.Length)
                    client.DownloadFile(remotePath, fileStream, AddressOf DownloadProgressCallback)
                End Using
                client.Disconnect()
            End Using

            If Not WaitForFileDownloadCompletion(localPath) Then
                LogMessage("File download did not complete in time.", "ERROR")
                WriteDebugLog("File download did not complete in silent mode.")
                If Not runSilently Then
                    Me.Invoke(Sub() lblStatus.Text = "Status: Download failed.")
                End If
                Return False
            End If

            If File.Exists(localPath) Then
                LogMessage($"Download completed successfully: {localPath}", "INFO")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  lblStatus.Text = "Status: Download completed."
                                  progressBar1.Value = 100
                              End Sub)
                End If
                Return True
            Else
                LogMessage("Download failed: File not found after transfer.", "ERROR")
                WriteDebugLog("Download failed: File not found after transfer in silent mode.")
                If Not runSilently Then
                    Me.Invoke(Sub() lblStatus.Text = "Status: Download failed.")
                End If
                Return False
            End If
        Catch ex As Exception
            LogMessage($"Error downloading update: {ex.Message}", "ERROR")
            WriteDebugLog($"Error downloading update in silent mode: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                Me.Invoke(Sub() lblStatus.Text = "Status: Download failed.")
            End If
            Return False
        End Try
    End Function

    Private Function WaitForFileDownloadCompletion(tempPath As String) As Boolean
        Dim maxWaitTime As Integer = 30
        Dim elapsed As Integer = 0
        LogMessage("Waiting for download to finish...", "INFO")

        While elapsed < maxWaitTime
            Try
                Using stream As FileStream = File.Open(tempPath, FileMode.Open, FileAccess.Read, FileShare.None)
                    LogMessage("Download complete.", "INFO")
                    Return True
                End Using
            Catch ex As IOException
                Threading.Thread.Sleep(1000)
                elapsed += 1
                LogMessage($"Waiting for download to finish... ({elapsed}s)", "INFO")
            End Try
        End While

        LogMessage("File still locked after timeout.", "ERROR")
        If runSilently Then WriteDebugLog("File still locked after timeout in silent mode.")
        Return False
    End Function

    Private Sub DownloadProgressCallback(totalBytesDownloaded As ULong)
        Dim progress As Integer = CInt((totalBytesDownloaded * 100) / totalBytes)
        If Not runSilently Then
            Me.Invoke(Sub()
                          lblStatus.Text = $"Status: Downloading update... {progress}%"
                          progressBar1.Value = Math.Min(progress, 100)
                      End Sub)
        End If

        If Not downloadComplete AndAlso (progress <> lastLoggedProgress OrElse progress = 100) Then
            LogMessage($"Download progress: {progress}%", "INFO")
            lastLoggedProgress = progress
            If progress = 100 Then downloadComplete = True
        End If
    End Sub

    Private Sub EnsureUpdateFolderExists()
        Try
            Dim folderPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), INSTALL_BASE_PATH)
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath)
                LogMessage("Folder 'NewPatchInstall' created in ProgramData.", "INFO")
            End If
        Catch ex As Exception
            LogMessage($"Error creating folder: {ex.Message}", "ERROR")
            WriteDebugLog($"Error creating folder in silent mode: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                MessageBox.Show($"Error creating folder: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    Private Function FindVastExecutable() As String
        Dim potentialPaths As String() = {
            "Program Files (x86)\MAM Software\VAST\VAST.exe",
            "Program Files\MAM Software\VAST\VAST.exe"
        }

        For Each drive As DriveInfo In DriveInfo.GetDrives()
            If drive.IsReady Then
                For Each relativePath As String In potentialPaths
                    Dim fullPath As String = Path.Combine(drive.Name, relativePath)
                    If File.Exists(fullPath) Then
                        LogMessage($"VAST executable found at: {fullPath}", "INFO")
                        Return fullPath
                    End If
                Next
            End If
        Next

        LogMessage("No VAST.exe found.", "WARNING")
        Return ""
    End Function

    Private Function GetFileVersion(filePath As String) As String
        Try
            If Not String.IsNullOrEmpty(filePath) AndAlso File.Exists(filePath) Then
                Return FileVersionInfo.GetVersionInfo(filePath).ProductVersion
            End If
            Return "0.0.0"
        Catch ex As Exception
            LogMessage($"Error retrieving file version: {ex.Message}", "ERROR")
            Return "0.0.0"
        End Try
    End Function

    Private Function GetLatestVersionFromSFTP(username As String, password As String, basePrefix As String) As String
        Try
            Using client As New SftpClient(SftpHost, username, password)
                client.Connect()
                Dim files = client.ListDirectory(RemoteDirectory) _
                             .Where(Function(f) f.IsRegularFile AndAlso f.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))

                Dim candidates As New List(Of Version)
                For Each file In files
                    Dim nameWithoutExt = Path.GetFileNameWithoutExtension(file.Name)
                    Dim ver As Version = Nothing
                    If Version.TryParse(nameWithoutExt, ver) AndAlso $"{ver.Major}.{ver.Minor}" = basePrefix Then
                        candidates.Add(ver)
                    End If
                Next
                client.Disconnect()

                If candidates.Count = 0 Then
                    LogMessage($"⚠️ No matching version found on FTP for prefix: {basePrefix}", "WARNING")
                    Return "0.0.0"
                End If

                candidates.Sort()
                Dim latest = candidates.Last().ToString()
                LogMessage($"✅ Latest version available for {basePrefix}: {latest}", "INFO")
                Return latest
            End Using
        Catch ex As Exception
            LogMessage($"❌ SFTP version check error: {ex.Message}", "ERROR")
            WriteDebugLog($"SFTP error: {ex.Message}, Stacktrace: {ex.StackTrace}")
            Return "0.0.0"
        End Try
    End Function

    Private Sub GrantSystemFullControl(path As String, isDirectory As Boolean)
        Try
            If isDirectory Then
                Dim dirInfo As New DirectoryInfo(path)
                Dim security As DirectorySecurity = dirInfo.GetAccessControl()
                Dim systemSid As SecurityIdentifier = New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)
                Dim systemAccount As NTAccount = systemSid.Translate(GetType(NTAccount))
                Dim fullControlRule As FileSystemAccessRule = New FileSystemAccessRule(systemAccount, FileSystemRights.FullControl, InheritanceFlags.ObjectInherit Or InheritanceFlags.ContainerInherit, PropagationFlags.None, AccessControlType.Allow)
                security.AddAccessRule(fullControlRule)
                dirInfo.SetAccessControl(security)
                LogMessage($"Granted SYSTEM full control to directory: {path}", "INFO")
            Else
                Dim fileInfo As New FileInfo(path)
                Dim security As FileSecurity = fileInfo.GetAccessControl()
                Dim systemSid As SecurityIdentifier = New SecurityIdentifier(WellKnownSidType.LocalSystemSid, Nothing)
                Dim systemAccount As NTAccount = systemSid.Translate(GetType(NTAccount))
                Dim fullControlRule As FileSystemAccessRule = New FileSystemAccessRule(systemAccount, FileSystemRights.FullControl, InheritanceFlags.None, PropagationFlags.None, AccessControlType.Allow)
                security.AddAccessRule(fullControlRule)
                fileInfo.SetAccessControl(security)
                LogMessage($"Granted SYSTEM full control to file: {path}", "INFO")
            End If
        Catch ex As Exception
            LogMessage($"❌ Failed to grant SYSTEM full control to {path}: {ex.Message}", "ERROR")
            WriteDebugLog($"Failed to grant SYSTEM full control: {ex.Message}, Stacktrace: {ex.StackTrace}")
        End Try
    End Sub

    Private Sub WriteDebugLog(message As String)
        Try
            Dim debugLogPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater", "DebugLog.txt")
            Dim logDir As String = Path.GetDirectoryName(debugLogPath)
            If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)
            File.AppendAllText(debugLogPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [DEBUG] {message}{Environment.NewLine}")
        Catch ex As Exception
            Try
                Dim fallbackPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "VASTUpdater_DebugLog.txt")
                File.AppendAllText(fallbackPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [DEBUG] {message}{Environment.NewLine}")
            Catch
                ' Silent fail
            End Try
        End Try
    End Sub

    Private Sub LogMessage(message As String, level As String)
        Dim maxRetries As Integer = 3
        Dim retryDelay As Integer = 1000

        For attempt As Integer = 1 To maxRetries
            Try
                Dim logPath As String = GetLogPath()
                Dim logDir As String = Path.GetDirectoryName(logPath)
                If Not Directory.Exists(logDir) Then Directory.CreateDirectory(logDir)
                Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}{Environment.NewLine}"
                Using fs As New FileStream(logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite)
                    Using writer As New StreamWriter(fs) With {.AutoFlush = True}
                        writer.Write(logEntry)
                    End Using
                End Using
                Debug.WriteLine(logEntry)
                Exit For
            Catch ex As IOException When attempt < maxRetries
                WriteDebugLog($"Attempt {attempt} failed to log message (file contention): {ex.Message}, retrying in {retryDelay}ms...")
                Threading.Thread.Sleep(retryDelay)
            Catch ex As Exception
                WriteDebugLog($"Failed to log message: {ex.Message}, Stacktrace: {ex.StackTrace}")
                Try
                    Dim fallbackPath As String = Path.Combine(Path.GetDirectoryName(GetLogPath()), $"UpdateLog_{DateTime.Now:yyyyMMdd_HHmmss}.txt")
                    File.AppendAllText(fallbackPath, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}{Environment.NewLine}")
                Catch
                    ' Silent fail
                End Try
                Exit For
            End Try
        Next
    End Sub

    Private Function GetLogPath() As String
        Dim logPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), LOG_BASE_PATH, "UpdateLog.txt")
        Try
            Dim logDir As String = Path.GetDirectoryName(logPath)
            If Not Directory.Exists(logDir) Then
                Directory.CreateDirectory(logDir)
                LogMessage($"Created log directory: {logDir}", "INFO")
            End If
            If Not File.Exists(logPath) Then File.WriteAllText(logPath, $"UpdateLog.txt created at {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}")
        Catch ex As Exception
            WriteDebugLog($"Failed to ensure log path: {ex.Message}")
            Dim fallbackPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VASTUpdater", "UpdateLog.txt")
            Dim fallbackDir As String = Path.GetDirectoryName(fallbackPath)
            If Not Directory.Exists(fallbackDir) Then Directory.CreateDirectory(fallbackDir)
            If Not File.Exists(fallbackPath) Then File.WriteAllText(fallbackPath, $"Fallback UpdateLog.txt created at {DateTime.Now:yyyy-MM-dd HH:mm:ss}{Environment.NewLine}")
            Return fallbackPath
        End Try
        Return logPath
    End Function

    Private Function GetInstallPath(version As String) As String
        Dim basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), INSTALL_BASE_PATH)
        If Not Directory.Exists(basePath) Then Directory.CreateDirectory(basePath)
        Return Path.Combine(basePath, $"{version}.exe")
    End Function

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        MyBase.OnFormClosing(e)
        LogMessage("Application closing...", "INFO")
        Try
            If fadeTimer IsNot Nothing Then
                fadeTimer.Stop()
                fadeTimer.Dispose()
                fadeTimer = Nothing
            End If
        Catch ex As Exception
            LogMessage($"Error during form closing: {ex.Message}", "ERROR")
        Finally
            Environment.Exit(0)
        End Try
    End Sub

    Private Sub CreateScheduledTask()
        Try
            Dim exePath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater\VastAutoUpdater.exe")
            If Not File.Exists(exePath) Then
                LogMessage($"❌ VastAutoUpdater.exe not found at: {exePath}", "ERROR")
                If Not runSilently Then
                    MessageBox.Show("Cannot create scheduled task. VastAutoUpdater.exe not found.", "Task Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return
            End If

            Dim taskCommand As String = "schtasks /create /tn ""VASTAutoUpdate"" /tr """"C:\ProgramData\VASTUpdater\VastAutoUpdater.exe"" silent"" /sc weekly /d SUN /st 02:00 /ru SYSTEM /rl HIGHEST /f"
            LogMessage($"Creating scheduled task with command: {taskCommand}", "INFO")

            Dim processInfo As New ProcessStartInfo With {
                .FileName = "cmd.exe",
                .Arguments = "/C " & taskCommand,
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }

            Using proc As Process = Process.Start(processInfo)
                Dim output As String = proc.StandardOutput.ReadToEnd()
                Dim errorOutput As String = proc.StandardError.ReadToEnd()
                proc.WaitForExit()

                If proc.ExitCode = 0 Then
                    LogMessage("Scheduled Task created successfully.", "INFO")
                    LogMessage($"Task creation output: {output}", "INFO")
                Else
                    LogMessage($"Failed to create scheduled task. Exit Code: {proc.ExitCode} - Error: {errorOutput}", "WARNING")
                    If Not runSilently Then
                        MessageBox.Show($"Failed to create the scheduled task. Error: {errorOutput}", "Task Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                    WriteDebugLog($"Failed to create scheduled task: Exit Code {proc.ExitCode}, Error: {errorOutput}")
                End If
            End Using
        Catch ex As Exception
            LogMessage($"Error creating scheduled task: {ex.Message}", "WARNING")
            WriteDebugLog($"Error creating scheduled task: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                MessageBox.Show("Failed to create the scheduled task. Please check your permissions or run as administrator.", "Task Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End Try
    End Sub

    Private Sub VASTUpdater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DisplayCurrentVersion()
    End Sub

    Private Sub DisplayCurrentVersion()
        Try
            Dim vastPath As String = FindVastExecutable()
            If String.IsNullOrEmpty(vastPath) Then
                LogMessage("VAST.exe Not Found", "WARNING")
                If Not runSilently Then
                    Me.Invoke(Sub()
                                  lblCurrentVersion.Text = "VAST.exe Not Found"
                                  lblStatus.Text = "Status: VAST.exe not found."
                              End Sub)
                End If
                Return
            End If

            Dim currentVersion As String = GetFileVersion(vastPath)
            LogMessage($"Current Version: {currentVersion}", "INFO")
            If Not runSilently Then
                Me.Invoke(Sub()
                              lblCurrentVersion.Text = $"Current Version: {currentVersion}"
                              lblStatus.Text = "Status: Ready for update check..."
                          End Sub)
            End If
        Catch ex As Exception
            LogMessage($"Current Version: Unknown - {ex.Message}", "ERROR")
            WriteDebugLog($"Display version failed: {ex.Message}, Stacktrace: {ex.StackTrace}")
            If Not runSilently Then
                Me.Invoke(Sub()
                              lblCurrentVersion.Text = "Current Version: Unknown"
                              lblStatus.Text = "Status: Error checking version."
                          End Sub)
            End If
        End Try
    End Sub

End Class

Imports System.IO
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Threading

''' <summary>
''' Core engine orchestrating update workflow.
''' Uses async retry, single SFTP connection, percentage-based progress,
''' and installer integrity verification.
''' </summary>
Public Class UpdaterEngine
    Private ReadOnly _emailService As IEmailService
    Private ReadOnly _sftpFactory As Func(Of ISftpService)
    Private ReadOnly _dashboardService As IDashboardService
    Private Const MAX_RETRIES As Integer = 3
    Private Const RETRY_DELAY_MS As Integer = 5000
    Private Const INSTALLER_WAIT_TIMEOUT_MS As Integer = 30 * 60 * 1000
    Private Const INSTALLER_POLL_MS As Integer = 5000

    ''' <summary>
    ''' Create an UpdaterEngine with default production services.
    ''' </summary>
    Public Sub New()
        _emailService = New EmailService()
        _sftpFactory = Function() New SftpService()
        _dashboardService = New DashboardService()
    End Sub

    ''' <summary>
    ''' Create an UpdaterEngine with injected dependencies for testing.
    ''' </summary>
    Public Sub New(emailService As IEmailService, sftpFactory As Func(Of ISftpService),
                   Optional dashboardService As IDashboardService = Nothing)
        _emailService = emailService
        _sftpFactory = sftpFactory
        _dashboardService = If(dashboardService, New DashboardService())
    End Sub

    ''' <summary>
    ''' Perform the update workflow with retry logic and send a summary email.
    ''' </summary>
    Public Async Function PerformUpdateCheck(username As String, password As String, progress As Action(Of Integer, String), Optional cancelToken As CancellationToken = Nothing) As Task
        Dim success As Boolean = False
        Dim message As String = String.Empty
        Dim caughtEx As Exception = Nothing
        Dim currentVersion As String = String.Empty
        Dim targetVersion As String = String.Empty
        Dim updateStarted As Boolean = False
        Dim verifiedVersion As String = String.Empty

        ' Honor cancellation before doing any work (drive scan, SFTP connect)
        cancelToken.ThrowIfCancellationRequested()

        Using sftp As ISftpService = _sftpFactory()
            Try
                Logger.Log("Running update check", Logger.LogLevel.Info)
                progress(5, "Locating VAST installation...")

                Dim vastPath As String = VersionService.FindVastExecutable()
                If String.IsNullOrEmpty(vastPath) Then
                    Throw New UpdateException(UpdateErrorCode.VastNotFound, "VAST.exe not found on any drive")
                End If

                currentVersion = VersionService.GetFileVersion(vastPath)
                Logger.Log($"Current VAST version: {currentVersion}", Logger.LogLevel.Info)

                Dim parsedCurrent As Version = Nothing
                If Not Version.TryParse(currentVersion, parsedCurrent) Then
                    Throw New UpdateException(UpdateErrorCode.VersionParseError, $"Cannot parse current version: {currentVersion}")
                End If
                Dim prefix As String = $"{parsedCurrent.Major}.{parsedCurrent.Minor}"

                cancelToken.ThrowIfCancellationRequested()
                progress(10, "Connecting to update server...")

                ' Connect once, reuse for version check + download
                Await RetryOperationAsync(
                    Async Function()
                        Await Task.Run(Sub() sftp.Connect(username, password))
                        Return True
                    End Function,
                    "SFTP connect")

                progress(15, "Checking for updates...")

                Dim latest As String = Await Task.Run(Function() sftp.GetLatestVersion(prefix))

                If latest = "0.0.0" Then
                    message = "No update available"
                    progress(100, "No update available.")
                    Logger.Log(message, Logger.LogLevel.Info)
                    success = True
                    Return
                End If

                Dim parsedLatest As Version = Nothing
                If Not Version.TryParse(latest, parsedLatest) Then
                    Throw New UpdateException(UpdateErrorCode.VersionParseError, $"Invalid remote version format: {latest}")
                End If

                If parsedLatest.CompareTo(parsedCurrent) <= 0 Then
                    message = $"Already up-to-date (current: {currentVersion}, remote: {latest})"
                    progress(100, "Already up-to-date.")
                    Logger.Log(message, Logger.LogLevel.Info)
                    success = True
                    Return
                End If

                cancelToken.ThrowIfCancellationRequested()
                Logger.Log($"Update available: {currentVersion} -> {latest}", Logger.LogLevel.Info)

                ' Report update start to the dashboard (fire-and-forget)
                targetVersion = latest
                updateStarted = True
                _dashboardService.ReportEvent(DashboardEventType.UpdateStart, currentVersion, targetVersion)

                progress(20, $"Downloading version {latest}...")

                InstallerPathService.EnsureUpdateFolderExists()
                Dim installer As String = InstallerPathService.GetInstallPath(latest)

                ' Get remote file size for accurate progress
                Dim remoteSize As Long = Await Task.Run(Function() sftp.GetRemoteFileSize($"{latest}.exe"))
                If remoteSize <= 0 Then
                    Throw New UpdateException(UpdateErrorCode.DownloadFailed, $"Remote installer file not found: {latest}.exe")
                End If

                Dim totalMb As Long = CLng(Math.Ceiling(remoteSize / 1048576.0))
                Dim downloadOk As Boolean = Await Task.Run(
                    Function() sftp.DownloadFile($"{latest}.exe", installer,
                        Sub(bytesTransferred As ULong)
                            ' Scale download progress from 20% to 90%
                            Dim ratio As Double = Math.Min(CDbl(bytesTransferred) / CDbl(remoteSize), 1.0)
                            Dim pct As Integer = 20 + CInt(Math.Floor(ratio * 70))
                            Dim mbTransferred As Long = CLng(bytesTransferred \ 1048576UL)
                            progress(pct, $"Downloading... {mbTransferred} MB of {totalMb} MB")
                        End Sub))

                If Not downloadOk Then
                    Throw New UpdateException(UpdateErrorCode.DownloadFailed, "Download failed — file is empty or missing")
                End If

                ' Verify downloaded file exists and has content
                Dim installerInfo As New FileInfo(installer)
                If Not installerInfo.Exists OrElse installerInfo.Length = 0 Then
                    Throw New UpdateException(UpdateErrorCode.DownloadFailed, "Downloaded installer is empty or missing")
                End If

                cancelToken.ThrowIfCancellationRequested()

                ' Verify installer integrity via SHA-256 hash sidecar (if available)
                Await VerifyInstallerHash(sftp, latest, installer)

                Logger.Log($"Download verified: {installer} ({installerInfo.Length} bytes)", Logger.LogLevel.Info)

                ' VAST programs must not be running while the patch installs.
                ' Applies to interactive and silent runs alike.
                Dim openPrograms As List(Of String) = VastProcessService.GetRunningVastPrograms()
                If openPrograms.Count > 0 Then
                    progress(93, $"Closing VAST programs ({String.Join(", ", openPrograms)})...")
                    Dim allClosed As Boolean = Await Task.Run(Function() VastProcessService.CloseVastPrograms())
                    If Not allClosed Then
                        Throw New UpdateException(UpdateErrorCode.ProcessCloseFailed,
                            "Could not close all running VAST programs before the update")
                    End If
                End If

                progress(95, "Launching installer...")

                ' Launch installer unattended. VAST patch installers accept /silent
                ' for a no-UI install, which also keeps the 2:00 AM SYSTEM
                ' scheduled task fully headless.
                Dim proc As Process = Nothing
                Dim installerExited As Boolean = False
                Try
                    Logger.Log($"Launching installer: ""{installer}"" /silent", Logger.LogLevel.Info)
                    proc = Process.Start(New ProcessStartInfo With {
                        .FileName = installer,
                        .Arguments = "/silent",
                        .UseShellExecute = True
                    })

                    If proc IsNot Nothing Then
                        ' Wait for the patch to finish (Setup64/msiexec can take
                        ' many minutes). Cancellation stops the wait, not the patch.
                        progress(96, "Installing patch...")
                        Dim waitedMs As Integer = 0
                        While waitedMs < INSTALLER_WAIT_TIMEOUT_MS
                            If cancelToken.IsCancellationRequested Then Exit While
                            installerExited = Await Task.Run(Function() proc.WaitForExit(INSTALLER_POLL_MS))
                            If installerExited Then Exit While
                            waitedMs += INSTALLER_POLL_MS
                            Dim mins As Integer = waitedMs \ 60000
                            progress(97, If(mins > 0, $"Installing patch... ({mins} min elapsed)", "Installing patch..."))
                        End While

                        If installerExited AndAlso proc.ExitCode <> 0 Then
                            Throw New UpdateException(UpdateErrorCode.InstallerFailed, $"Installer exited with error code {proc.ExitCode}")
                        End If
                    End If
                Catch ex As UpdateException
                    Throw
                Catch ex As Exception
                    Throw New UpdateException(UpdateErrorCode.InstallerFailed, $"Failed to launch installer: {ex.Message}", ex)
                Finally
                    If proc IsNot Nothing Then proc.Dispose()
                End Try

                ' Clean up old installers (never touches the current version's file)
                InstallerPathService.CleanupOldInstallers(latest)

                ' Setup has closed = the patch is done: read the on-disk VAST
                ' version and report exactly that. Only a wait timeout leaves
                ' the outcome unconfirmed.
                success = True
                If installerExited Then
                    Dim diskVersion As String = GetInstalledVersionString()
                    verifiedVersion = If(String.IsNullOrEmpty(diskVersion), currentVersion, diskVersion)

                    Dim parsedDisk As Version = Nothing
                    If Version.TryParse(diskVersion, parsedDisk) AndAlso parsedDisk.CompareTo(parsedLatest) >= 0 Then
                        message = $"Update {latest} installed — VAST is now {diskVersion}"
                        Logger.Log(message, Logger.LogLevel.Info)
                        progress(100, "Update complete.")
                    Else
                        message = $"Installer finished — VAST reports version {If(String.IsNullOrEmpty(diskVersion), "unknown", diskVersion)}"
                        Logger.Log(message, Logger.LogLevel.Warning)
                        progress(100, message)
                    End If
                Else
                    message = $"Update {latest} is still installing in the background"
                    Logger.Log(message, Logger.LogLevel.Info)
                    progress(100, "Patch is installing in the background...")
                End If

            Catch ex As OperationCanceledException
                ' Let cancellation propagate directly — do not wrap or email
                Throw
            Catch ex As Exception
                caughtEx = ex
                Logger.Log($"Update error: {ex.Message}", Logger.LogLevel.Error)
                If ex.StackTrace IsNot Nothing Then
                    Logger.Log($"Stack trace: {ex.StackTrace}", Logger.LogLevel.Error)
                End If
                message = ex.Message
            End Try
        End Using ' SftpService disconnects and disposes here

        ' Report outcome to the dashboard (fire-and-forget, never blocks the update)
        Try
            If caughtEx IsNot Nothing Then
                _dashboardService.ReportEvent(DashboardEventType.UpdateFailure, currentVersion, targetVersion, message)
            ElseIf updateStarted AndAlso success AndAlso verifiedVersion <> String.Empty Then
                ' Setup has closed — report the version actually on disk so the
                ' dashboard reflects the machine's real state
                _dashboardService.ReportEvent(DashboardEventType.UpdateSuccess, verifiedVersion, targetVersion, message)
            ElseIf updateStarted AndAlso success Then
                ' Wait timed out with setup still running — leave update_start
                ' as the machine's latest state; the next run's heartbeat
                ' reports the version once the patch has landed
                Logger.Log("Skipping update_success report — installer still running", Logger.LogLevel.Info)
            End If
        Catch dashEx As Exception
            Logger.Log($"Failed to report dashboard status: {dashEx.Message}", Logger.LogLevel.Warning)
        End Try

        ' Send summary email
        Try
            Await Task.Run(Sub() _emailService.SendSummary(success, message, caughtEx))
        Catch emailEx As Exception
            Logger.Log($"Failed to send summary email: {emailEx.Message}", Logger.LogLevel.Warning)
        End Try

        ' Re-throw so the caller can display the error
        If caughtEx IsNot Nothing Then
            If TypeOf caughtEx Is UpdateException Then
                Throw caughtEx
            End If
            Throw New UpdateException(UpdateErrorCode.Unknown, message, caughtEx)
        End If
    End Function

    ''' <summary>
    ''' Read the currently installed VAST version from disk.
    ''' Returns an empty string when it cannot be determined.
    ''' </summary>
    Private Shared Function GetInstalledVersionString() As String
        Try
            Dim exePath As String = VersionService.FindVastExecutable()
            If String.IsNullOrEmpty(exePath) Then Return String.Empty
            Return VersionService.GetFileVersion(exePath)
        Catch ex As Exception
            Logger.Log($"Post-install version check failed: {ex.Message}", Logger.LogLevel.Warning)
            Return String.Empty
        End Try
    End Function

    ''' <summary>
    ''' Verify installer integrity using a SHA-256 hash sidecar file (.sha256).
    ''' If no sidecar exists on the server, logs a warning and continues.
    ''' </summary>
    Private Async Function VerifyInstallerHash(sftp As ISftpService, version As String, localPath As String) As Task
        Try
            Dim hashFileName As String = $"{version}.exe.sha256"
            Dim hashSize As Long = Await Task.Run(Function() sftp.GetRemoteFileSize(hashFileName))
            If hashSize <= 0 Then
                Logger.Log("No SHA-256 hash file available on server — skipping integrity check", Logger.LogLevel.Warning)
                Return
            End If

            ' Download the hash file
            Dim hashPath As String = localPath & ".sha256"
            Await Task.Run(Function() sftp.DownloadFile(hashFileName, hashPath, Sub(b) Return))

            ' Read expected hash
            Dim expectedHash As String = File.ReadAllText(hashPath).Trim().Split(" "c)(0).ToUpperInvariant()

            ' Compute actual hash
            Dim actualHash As String
            Using sha As SHA256 = SHA256.Create()
                Using fs As New FileStream(localPath, FileMode.Open, FileAccess.Read, FileShare.Read)
                    Dim hashBytes As Byte() = sha.ComputeHash(fs)
                    actualHash = BitConverter.ToString(hashBytes).Replace("-", "").ToUpperInvariant()
                End Using
            End Using

            If actualHash <> expectedHash Then
                ' Delete the corrupt installer
                Try
                    File.Delete(localPath)
                Catch
                End Try
                Throw New UpdateException(UpdateErrorCode.HashMismatch, $"Installer hash mismatch! Expected: {expectedHash}, Got: {actualHash}")
            End If

            Logger.Log("Installer integrity verified via SHA-256", Logger.LogLevel.Info)

            ' Clean up hash file
            Try
                File.Delete(hashPath)
            Catch
            End Try
        Catch ex As UpdateException
            Throw ' Re-throw hash mismatch
        Catch ex As Exception
            Logger.Log($"Hash verification skipped: {ex.Message}", Logger.LogLevel.Warning)
        End Try
    End Function

    ''' <summary>
    ''' Retry an async operation up to MAX_RETRIES times with Task.Delay between attempts.
    ''' Does not block thread pool threads.
    ''' </summary>
    Private Async Function RetryOperationAsync(Of T)(operation As Func(Of Task(Of T)), operationName As String) As Task(Of T)
        Dim lastEx As Exception = Nothing
        For attempt As Integer = 1 To MAX_RETRIES
            Dim shouldDelay As Boolean = False
            Try
                Return Await operation()
            Catch ex As Exception
                lastEx = ex
                Logger.Log($"Attempt {attempt}/{MAX_RETRIES} for {operationName} failed: {ex.Message}", Logger.LogLevel.Warning)
                shouldDelay = (attempt < MAX_RETRIES)
            End Try
            ' Await cannot be inside Catch in VB.NET (.NET Framework)
            If shouldDelay Then
                Await Task.Delay(RETRY_DELAY_MS * attempt)
            End If
        Next
        Throw New UpdateException(UpdateErrorCode.ConnectionFailed, $"{operationName} failed after {MAX_RETRIES} attempts", lastEx)
    End Function

End Class

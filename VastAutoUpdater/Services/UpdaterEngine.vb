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
    Private Const MAX_RETRIES As Integer = 3
    Private Const RETRY_DELAY_MS As Integer = 5000

    ''' <summary>
    ''' Create an UpdaterEngine with default production services.
    ''' </summary>
    Public Sub New()
        _emailService = New EmailService()
        _sftpFactory = Function() New SftpService()
    End Sub

    ''' <summary>
    ''' Create an UpdaterEngine with injected dependencies for testing.
    ''' </summary>
    Public Sub New(emailService As IEmailService, sftpFactory As Func(Of ISftpService))
        _emailService = emailService
        _sftpFactory = sftpFactory
    End Sub

    ''' <summary>
    ''' Perform the update workflow with retry logic and send a summary email.
    ''' </summary>
    Public Async Function PerformUpdateCheck(username As String, password As String, progress As Action(Of Integer, String), Optional cancelToken As CancellationToken = Nothing) As Task
        Dim success As Boolean = False
        Dim message As String = String.Empty
        Dim caughtEx As Exception = Nothing

        Using sftp As ISftpService = _sftpFactory()
            Try
                Logger.Log("Running update check", Logger.LogLevel.Info)
                progress(5, "Locating VAST installation...")

                Dim vastPath As String = VersionService.FindVastExecutable()
                If String.IsNullOrEmpty(vastPath) Then
                    Throw New UpdateException(UpdateErrorCode.VastNotFound, "VAST.exe not found on any drive")
                End If

                Dim currentVersion As String = VersionService.GetFileVersion(vastPath)
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
                progress(20, $"Downloading version {latest}...")

                InstallerPathService.EnsureUpdateFolderExists()
                Dim installer As String = InstallerPathService.GetInstallPath(latest)

                ' Get remote file size for accurate progress
                Dim remoteSize As Long = Await Task.Run(Function() sftp.GetRemoteFileSize($"{latest}.exe"))
                If remoteSize <= 0 Then
                    Throw New UpdateException(UpdateErrorCode.DownloadFailed, $"Remote installer file not found: {latest}.exe")
                End If

                Dim downloadOk As Boolean = Await Task.Run(
                    Function() sftp.DownloadFile($"{latest}.exe", installer,
                        Sub(bytesTransferred As ULong)
                            ' Scale download progress from 20% to 90%
                            Dim ratio As Double = Math.Min(CDbl(bytesTransferred) / CDbl(remoteSize), 1.0)
                            Dim pct As Integer = 20 + CInt(Math.Floor(ratio * 70))
                            Dim mbTransferred As Long = CLng(bytesTransferred \ 1048576UL)
                            progress(pct, $"Downloading... {mbTransferred} MB")
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
                progress(95, "Launching installer...")

                ' Launch installer and wait briefly to confirm it started
                Dim proc As Process = Nothing
                Try
                    proc = Process.Start(New ProcessStartInfo With {
                        .FileName = installer,
                        .UseShellExecute = True
                    })

                    If proc IsNot Nothing Then
                        Dim exited As Boolean = Await Task.Run(Function() proc.WaitForExit(10000))
                        If exited AndAlso proc.ExitCode <> 0 Then
                            Throw New UpdateException(UpdateErrorCode.InstallerFailed, $"Installer exited with error code {proc.ExitCode}")
                        ElseIf exited Then
                            success = True
                            message = $"Update {latest} downloaded and installer completed"
                            Logger.Log($"Installer completed: {installer} (exit code 0)", Logger.LogLevel.Info)
                        Else
                            success = True
                            message = $"Update {latest} downloaded and installer launched"
                            Logger.Log($"Installer still running: {installer} (PID: {proc.Id})", Logger.LogLevel.Info)
                        End If
                    End If
                Catch ex As Exception
                    Throw New UpdateException(UpdateErrorCode.InstallerFailed, $"Failed to launch installer: {ex.Message}", ex)
                Finally
                    If proc IsNot Nothing Then proc.Dispose()
                End Try

                ' Clean up old installers
                InstallerPathService.CleanupOldInstallers(latest)

                progress(100, "Update complete.")

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

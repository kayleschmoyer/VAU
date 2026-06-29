''' <summary>
''' Core engine orchestrating update workflow.
''' Includes retry logic, proper progress calculation, and installer verification.
''' </summary>
Imports System.IO
Imports System.Diagnostics
Imports System.Security.Cryptography

Public Class UpdaterEngine
    Private ReadOnly sftp As New SftpService()
    Private ReadOnly email As New EmailService()
    Private Const MAX_RETRIES As Integer = 3
    Private Const RETRY_DELAY_MS As Integer = 5000

    ''' <summary>
    ''' Perform the update workflow with retry logic and send a summary email.
    ''' </summary>
    Public Async Function PerformUpdateCheck(username As String, password As String, progress As Action(Of Integer, String)) As Task
        Dim success As Boolean = False
        Dim message As String = String.Empty
        Dim caughtEx As Exception = Nothing

        Try
            Logger.Log("Running update check", Logger.LogLevel.Info)
            progress(5, "Locating VAST installation...")

            Dim vastPath As String = VersionService.FindVastExecutable()
            If String.IsNullOrEmpty(vastPath) Then
                Throw New FileNotFoundException("VAST.exe not found on any drive")
            End If

            Dim currentVersion As String = VersionService.GetFileVersion(vastPath)
            Logger.Log($"Current VAST version: {currentVersion}", Logger.LogLevel.Info)

            Dim parsedCurrent As Version = Nothing
            If Not Version.TryParse(currentVersion, parsedCurrent) Then
                Throw New FormatException($"Cannot parse current version: {currentVersion}")
            End If
            Dim prefix As String = $"{parsedCurrent.Major}.{parsedCurrent.Minor}"

            progress(15, "Checking for updates...")

            ' Retry wrapper for version check
            Dim latest As String = Await Task.Run(Function() RetryOperation(
                Function() sftp.GetLatestVersion(username, password, prefix),
                "version check"))

            If latest = "0.0.0" Then
                message = "No update available"
                progress(100, "No update available.")
                Logger.Log(message, Logger.LogLevel.Info)
                success = True
                Return
            End If

            Dim parsedLatest As Version = Nothing
            If Not Version.TryParse(latest, parsedLatest) Then
                message = $"Invalid remote version format: {latest}"
                Logger.Log(message, Logger.LogLevel.Warning)
                Return
            End If

            If parsedLatest.CompareTo(parsedCurrent) <= 0 Then
                message = $"Already up-to-date (current: {currentVersion}, remote: {latest})"
                progress(100, "Already up-to-date.")
                Logger.Log(message, Logger.LogLevel.Info)
                success = True
                Return
            End If

            Logger.Log($"Update available: {currentVersion} -> {latest}", Logger.LogLevel.Info)
            progress(25, $"Downloading version {latest}...")

            InstallerPathService.EnsureUpdateFolderExists()
            Dim installer As String = InstallerPathService.GetInstallPath(latest)

            ' Calculate progress percentage based on file size
            Dim downloadOk As Boolean = Await Task.Run(Function() RetryOperation(
                Function() sftp.DownloadFile(username, password, latest, installer,
                    Sub(bytesTransferred As ULong)
                        ' Clamp to safe integer range for progress (25-90% of bar)
                        Dim pct As Integer = 25
                        If bytesTransferred > 0UL Then
                            ' Cap at 90% during download, leave 10% for install
                            pct = Math.Min(CInt(Math.Min(bytesTransferred \ 1048576UL, 65UL)) + 25, 90)
                        End If
                        progress(pct, $"Downloading... {bytesTransferred \ 1048576UL} MB")
                    End Sub),
                "file download"))

            If Not downloadOk Then
                message = "Download failed after retries"
                Logger.Log(message, Logger.LogLevel.Error)
                Return
            End If

            ' Verify downloaded file exists and has content
            Dim installerInfo As New FileInfo(installer)
            If Not installerInfo.Exists OrElse installerInfo.Length = 0 Then
                message = "Downloaded installer is empty or missing"
                Logger.Log(message, Logger.LogLevel.Error)
                Return
            End If

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
                    ' Wait up to 10 seconds to confirm the process started
                    Await Task.Run(Sub() proc.WaitForExit(10000))
                    Logger.Log($"Installer launched: {installer} (PID: {proc.Id})", Logger.LogLevel.Info)
                End If

                success = True
                message = $"Update {latest} downloaded and installer launched"
            Catch ex As Exception
                message = $"Failed to launch installer: {ex.Message}"
                Logger.Log(message, Logger.LogLevel.Error)
            Finally
                If proc IsNot Nothing Then proc.Dispose()
            End Try

            ' Clean up old installers
            InstallerPathService.CleanupOldInstallers(latest)

            progress(100, "Update complete.")

        Catch ex As Exception
            caughtEx = ex
            Logger.Log($"Update error: {ex.Message}", Logger.LogLevel.Error)
            If ex.StackTrace IsNot Nothing Then
                Logger.Log($"Stack trace: {ex.StackTrace}", Logger.LogLevel.Error)
            End If
            message = ex.Message
        Finally
            Try
                Await Task.Run(Sub() email.SendSummary(success, message, caughtEx))
            Catch emailEx As Exception
                Logger.Log($"Failed to send summary email: {emailEx.Message}", Logger.LogLevel.Warning)
            End Try
        End Try
    End Function

    ''' <summary>
    ''' Retry an operation up to MAX_RETRIES times with delay between attempts.
    ''' </summary>
    Private Function RetryOperation(Of T)(operation As Func(Of T), operationName As String) As T
        Dim lastEx As Exception = Nothing
        For attempt As Integer = 1 To MAX_RETRIES
            Try
                Return operation()
            Catch ex As Exception
                lastEx = ex
                Logger.Log($"Attempt {attempt}/{MAX_RETRIES} for {operationName} failed: {ex.Message}", Logger.LogLevel.Warning)
                If attempt < MAX_RETRIES Then
                    System.Threading.Thread.Sleep(RETRY_DELAY_MS * attempt)
                End If
            End Try
        Next
        Throw New Exception($"{operationName} failed after {MAX_RETRIES} attempts", lastEx)
    End Function

End Class

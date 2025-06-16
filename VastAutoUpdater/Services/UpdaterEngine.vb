''' <summary>
''' Core engine orchestrating update workflow.
''' </summary>
Imports System.IO
Imports System.Diagnostics

Public Class UpdaterEngine
    Private ReadOnly sftp As New SftpService()
    Private ReadOnly email As New EmailService()

    ''' <summary>
    ''' Perform the update workflow and send a summary email upon completion.
    ''' </summary>
    Public Async Function PerformUpdateCheck(username As String, password As String, progress As Action(Of Integer, String)) As Task
        Dim success As Boolean = False
        Dim message As String = String.Empty
        Dim caughtEx As Exception = Nothing

        Try
            Logger.Log("Running update check", Logger.LogLevel.Info)

            Dim vastPath = VersionService.FindVastExecutable()
            If String.IsNullOrEmpty(vastPath) Then
                Throw New FileNotFoundException("VAST.exe not found")
            End If

            Dim currentVersion = VersionService.GetFileVersion(vastPath)
            Dim prefix = $"{New Version(currentVersion).Major}.{New Version(currentVersion).Minor}"

            ' Wrap blocking network operation in Task.Run
            Dim latest As String = Await Task.Run(Function() sftp.GetLatestVersion(username, password, prefix))

            If latest = "0.0.0" Then
                message = "No update available"
                Return
            End If
            If New Version(latest).CompareTo(New Version(currentVersion)) <= 0 Then
                message = "Already up-to-date"
                Return
            End If

            InstallerPathService.EnsureUpdateFolderExists()
            Dim installer = InstallerPathService.GetInstallPath(latest)

            ' Wrap blocking file download in Task.Run
            Dim ok As Boolean = Await Task.Run(Function() sftp.DownloadFile(username, password, latest, installer, Sub(b) progress(CInt(b), "Downloading")))

            If ok Then
                Process.Start(installer)
                success = True
                message = $"Installed {latest}"
            Else
                message = "Download failed"
            End If
        Catch ex As Exception
            caughtEx = ex
            Logger.Log($"Update error: {ex.Message}", Logger.LogLevel.Error)
            message = ex.Message
        Finally
            Await Task.Run(Sub() email.SendSummary(success, message, caughtEx))
        End Try
    End Function

End Class

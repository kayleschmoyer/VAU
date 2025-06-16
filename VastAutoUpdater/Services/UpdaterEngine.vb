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

            Dim vastPath = FindVastExecutable()
            If String.IsNullOrEmpty(vastPath) Then
                Throw New FileNotFoundException("VAST.exe not found")
            End If

            Dim currentVersion = GetFileVersion(vastPath)
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

            EnsureUpdateFolderExists()
            Dim installer = GetInstallPath(latest)

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

    Public Shared Function FindVastExecutable() As String
        Dim potential As String() = {
            "Program Files (x86)\MAM Software\VAST\VAST.exe",
            "Program Files\MAM Software\VAST\VAST.exe"
        }
        For Each drive In DriveInfo.GetDrives()
            If drive.IsReady Then
                For Each rel In potential
                    Dim full = Path.Combine(drive.Name, rel)
                    If File.Exists(full) Then
                        Logger.Log($"VAST executable found at: {full}", Logger.LogLevel.Info)
                        Return full
                    End If
                Next
            End If
        Next
        Logger.Log("No VAST.exe found", Logger.LogLevel.Warning)
        Return String.Empty
    End Function

    Public Shared Function GetFileVersion(filePath As String) As String
        Try
            If File.Exists(filePath) Then
                Return FileVersionInfo.GetVersionInfo(filePath).ProductVersion
            End If
        Catch ex As Exception
            Logger.Log($"Error retrieving file version: {ex.Message}", Logger.LogLevel.Error)
        End Try
        Return "0.0.0"
    End Function

    Private Sub EnsureUpdateFolderExists()
        Dim folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater\NewPatchInstall")
        If Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If
    End Sub

    Private Function GetInstallPath(version As String) As String
        Dim basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "VASTUpdater\NewPatchInstall")
        If Not Directory.Exists(basePath) Then Directory.CreateDirectory(basePath)
        Return Path.Combine(basePath, $"{version}.exe")
    End Function
End Class

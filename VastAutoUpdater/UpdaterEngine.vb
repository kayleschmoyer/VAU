Imports System.IO
Imports System.Diagnostics

Public Class UpdaterEngine
    Private ReadOnly sftp As New SftpService()
    Private ReadOnly email As New EmailService()

    Public Async Function PerformUpdateCheck(username As String, password As String, progress As Action(Of Integer, String)) As Task
        Logger.Log("Running update check", Logger.LogLevel.Info)

        Dim vastPath = FindVastExecutable()
        If String.IsNullOrEmpty(vastPath) Then
            Logger.Log("VAST.exe not found", Logger.LogLevel.Error)
            Throw New FileNotFoundException("VAST.exe not found")
        End If

        Dim currentVersion = GetFileVersion(vastPath)
        Dim prefix = $"{New Version(currentVersion).Major}.{New Version(currentVersion).Minor}"

        ' Wrap blocking network operation in Task.Run
        Dim latest As String = Await Task.Run(Function() sftp.GetLatestVersion(username, password, prefix))

        If latest = "0.0.0" Then Return
        If New Version(latest).CompareTo(New Version(currentVersion)) <= 0 Then
            Logger.Log("Already up-to-date", Logger.LogLevel.Info)
            Return
        End If

        EnsureUpdateFolderExists()
        Dim installer = GetInstallPath(latest)

        ' Wrap blocking file download in Task.Run
        Dim ok As Boolean = Await Task.Run(Function() sftp.DownloadFile(username, password, latest, installer, Sub(b) progress(CInt(b), "Downloading")))

        If ok Then
            Process.Start(installer)
            Await Task.Run(Sub() email.SendNotification(True, $"Installed {latest}"))
        Else
            Await Task.Run(Sub() email.SendNotification(False, "Download failed"))
        End If
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

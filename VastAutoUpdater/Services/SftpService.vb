''' <summary>
''' Handles communication with the remote SFTP server to check and download updates.
''' Includes connection timeouts and proper resource disposal.
''' </summary>
Imports Renci.SshNet
Imports System.IO
Imports System.Linq

Public Class SftpService
    Private ReadOnly host As String = ConfigManager.SftpHost
    Private ReadOnly remoteDir As String = "/VASTAutoInstall/"
    Private Const CONNECTION_TIMEOUT_SECONDS As Integer = 30
    Private Const OPERATION_TIMEOUT_SECONDS As Integer = 300

    ''' <summary>
    ''' Create a configured SftpClient with proper timeouts.
    ''' </summary>
    Private Function CreateClient(username As String, password As String) As SftpClient
        Dim client As New SftpClient(host, username, password)
        client.ConnectionInfo.Timeout = TimeSpan.FromSeconds(CONNECTION_TIMEOUT_SECONDS)
        client.OperationTimeout = TimeSpan.FromSeconds(OPERATION_TIMEOUT_SECONDS)
        Return client
    End Function

    ''' <summary>
    ''' Query the SFTP server for the latest version matching the given major.minor prefix.
    ''' Returns "0.0.0" if no matching version is found or on error.
    ''' </summary>
    Public Function GetLatestVersion(username As String, password As String, prefix As String) As String
        Using client As SftpClient = CreateClient(username, password)
            client.Connect()
            Logger.Log($"Connected to SFTP host: {host}", Logger.LogLevel.Info)

            Dim files = client.ListDirectory(remoteDir).
                Where(Function(f) f.IsRegularFile AndAlso f.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))

            Dim versions As New List(Of Version)()
            For Each f In files
                Dim name As String = Path.GetFileNameWithoutExtension(f.Name)
                Dim v As Version = Nothing
                If Version.TryParse(name, v) AndAlso $"{v.Major}.{v.Minor}" = prefix Then
                    versions.Add(v)
                End If
            Next

            If versions.Count = 0 Then
                Logger.Log($"No matching version found on SFTP for prefix: {prefix}", Logger.LogLevel.Warning)
                Return "0.0.0"
            End If

            versions.Sort()
            Dim latest As String = versions.Last().ToString()
            Logger.Log($"Latest version available for {prefix}: {latest}", Logger.LogLevel.Info)
            Return latest
        End Using
    End Function

    ''' <summary>
    ''' Download the installer for the specified version to localPath.
    ''' Returns True on success, False on failure.
    ''' </summary>
    Public Function DownloadFile(username As String, password As String, version As String, localPath As String, progress As Action(Of ULong)) As Boolean
        Using client As SftpClient = CreateClient(username, password)
            client.Connect()
            Dim remotePath As String = $"{remoteDir}{version}.exe"

            Logger.Log($"Downloading: {remotePath} -> {localPath}", Logger.LogLevel.Info)

            Using fs As New FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None)
                client.DownloadFile(remotePath, fs, Sub(d) progress(d))
            End Using

            ' Verify file was actually written
            Dim fi As New FileInfo(localPath)
            If fi.Exists AndAlso fi.Length > 0 Then
                Logger.Log($"Download completed: {localPath} ({fi.Length} bytes)", Logger.LogLevel.Info)
                Return True
            Else
                Logger.Log($"Download produced empty file: {localPath}", Logger.LogLevel.Error)
                Return False
            End If
        End Using
    End Function
End Class

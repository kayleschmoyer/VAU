Imports Renci.SshNet

Public Class SftpService
    Private ReadOnly host As String = ConfigManager.SftpHost
    Private ReadOnly remoteDir As String = "/VASTAutoInstall/"

    Public Function GetLatestVersion(username As String, password As String, prefix As String) As String
        Try
            Using client As New SftpClient(host, username, password)
                client.Connect()
                Dim files = client.ListDirectory(remoteDir).
                    Where(Function(f) f.IsRegularFile AndAlso f.Name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                Dim versions As New List(Of Version)()
                For Each f In files
                    Dim name = Path.GetFileNameWithoutExtension(f.Name)
                    Dim v As Version = Nothing
                    If Version.TryParse(name, v) AndAlso $"{v.Major}.{v.Minor}" = prefix Then
                        versions.Add(v)
                    End If
                Next
                client.Disconnect()
                If versions.Count = 0 Then
                    Logger.Log($"No matching version found on FTP for prefix: {prefix}", Logger.LogLevel.Warning)
                    Return "0.0.0"
                End If
                versions.Sort()
                Dim latest = versions.Last().ToString()
                Logger.Log($"Latest version available for {prefix}: {latest}", Logger.LogLevel.Info)
                Return latest
            End Using
        Catch ex As Exception
            Logger.Log($"SFTP version check error: {ex.Message}", Logger.LogLevel.Error)
            Return "0.0.0"
        End Try
    End Function

    Public Function DownloadFile(username As String, password As String, version As String, localPath As String, progress As Action(Of ULong)) As Boolean
        Try
            Using client As New SftpClient(host, username, password)
                client.Connect()
                Dim remotePath As String = $"{remoteDir}{version}.exe"
                Using fs As New FileStream(localPath, FileMode.Create, FileAccess.Write)
                    Dim info = client.Get(remotePath)
                    client.DownloadFile(remotePath, fs, Sub(d) progress(d))
                End Using
                client.Disconnect()
            End Using
            Logger.Log($"Download completed successfully: {localPath}", Logger.LogLevel.Info)
            Return True
        Catch ex As Exception
            Logger.Log($"Error downloading update: {ex.Message}", Logger.LogLevel.Error)
            Return False
        End Try
    End Function
End Class
